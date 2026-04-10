/*
Titre : Moteur d'Orchestration Wispr.Bridge
Auteur : Digixtp
Version : 4.0.0
Note introductive : Ce script Go agit comme un pont d'orchestration pour le moteur de transcription Whisper (CUDA). Il scanne un répertoire d'entrée, exécute l'inférence audio en mode masqué (Headless) et génère une note Markdown enrichie sémantiquement avant de l'injecter dans Joplin via son API locale. Il garantit une sécurité absolue des données via une mise en quarantaine horodatée (Move-Only) et intègre une validation stricte de l'environnement (Fail-Fast).
Prérequis : Go 1.20+, Whisper.cpp (binaire CLI compilé CUDA), Modèle Large-V3-Turbo, Joplin (Web Clipper activé).
Commande d'utilisation : go run Wispr_Bridge.go
*/

package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

// -----------------------------------------------------------------------------
// Configuration globale, Variables d'environnement et Charte Graphique ANSI
// -----------------------------------------------------------------------------
const (
	JoplinToken = "034a2ee109ad401f8246296d7def3edc28dd73accdb49744f1438227784990e56d4e011cdbaa903282931f185dc5ebaee6e9fa5f85de2951612c4f7d0deac651"
	JoplinPort  = "41184"
	TargetName  = "Wispr_Bridge"
)

const (
	ANSI_Reset  = "\033[0m"
	ANSI_Vert   = "\033[92m" // Titre principal, note introductive, sous-parties, succès
	ANSI_Orange = "\033[93m" // Avertissements, info, alertes non bloquantes
	ANSI_Rouge  = "\033[91m" // Erreurs critiques et blocages système
)

// Structures JSON pour l'API Joplin
type JoplinNote struct {
	ParentID string `json:"parent_id"`
	Title    string `json:"title"`
	Body     string `json:"body"`
}

type JoplinFolder struct {
	ID    string `json:"id"`
	Title string `json:"title"`
}

type JoplinFoldersResponse struct {
	Items []JoplinFolder `json:"items"`
}

// Structure de la configuration sémantique auto-générée
type SemanticConfig struct {
	Note     string            `json:"_note_introductive"`
	Keywords map[string]string `json:"domaines_et_mots_cles"`
}

// Hook bas niveau pour l'interprétation native des couleurs ANSI sous Windows
func init() {
	handle := syscall.Handle(os.Stdout.Fd())
	var mode uint32
	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	getConsoleMode := kernel32.NewProc("GetConsoleMode")
	setConsoleMode := kernel32.NewProc("SetConsoleMode")
	getConsoleMode.Call(uintptr(handle), uintptr(unsafe.Pointer(&mode)))
	mode |= 0x0004 // ENABLE_VIRTUAL_TERMINAL_PROCESSING
	setConsoleMode.Call(uintptr(handle), uintptr(mode))
}

func main() {
	fmt.Printf("%s=== Moteur Wispr.Bridge v4.0.0 (Production) ===%s\n", ANSI_Vert, ANSI_Reset)

	// -----------------------------------------------------------------------------
	// Phase Préliminaire : Validation des prérequis (Fail-Fast) et Routage dynamique
	// -----------------------------------------------------------------------------
	notebookID := getJoplinFolderID(TargetName)
	if notebookID == "" {
		fmt.Printf("%s[ERREUR] Carnet '%s' introuvable dans Joplin. Veuillez le créer et relancer.%s\n", ANSI_Rouge, TargetName, ANSI_Reset)
		return
	}

	// Définition dynamique des chemins (Agnosticisme matériel)
	homeDir, _ := os.UserHomeDir()
	inputDir := filepath.Join(homeDir, "Documents", "Wispr_Bridge")
	quarantineDir := filepath.Join(homeDir, "Desktop", "fichier à supprimer")
	modelsDir := filepath.Join(homeDir, "Documents", "whisper.cpp", "models")
	jsonConfigPath := filepath.Join(inputDir, "semantic_config.json")

	os.MkdirAll(inputDir, 0755)
	os.MkdirAll(quarantineDir, 0755)

	// Chargement du dictionnaire sémantique dynamique
	semanticRules := loadOrGenerateSemanticConfig(jsonConfigPath)

	// Découverte du binaire d'inférence
	whisperCliPath := findExecutable(homeDir)
	if whisperCliPath == "" {
		fmt.Printf("%s[ERREUR] Exécutable whisper-cli.exe introuvable.%s\n", ANSI_Rouge, ANSI_Reset)
		return
	}

	// Découverte du modèle IA Large-V3-Turbo
	whisperModelPath := findModel(modelsDir)
	if whisperModelPath == "" {
		fmt.Printf("%s[ERREUR] Aucun modèle IA (large-v3-turbo) trouvé dans le dossier models.%s\n", ANSI_Rouge, ANSI_Reset)
		return
	}

	fmt.Printf("%s[OK] Environnement validé. Modèle : %s%s\n", ANSI_Vert, filepath.Base(whisperModelPath), ANSI_Reset)

	// -----------------------------------------------------------------------------
	// Phase 1 : Scan unique du dossier (Exécution Batch)
	// -----------------------------------------------------------------------------
	audioFiles := scanDirectoryForAudio(inputDir)
	if len(audioFiles) == 0 {
		fmt.Printf("\n%s[INFO] Aucun fichier audio à traiter dans l'entrée. Arrêt propre du moteur.%s\n", ANSI_Orange, ANSI_Reset)
		time.Sleep(2 * time.Second)
		return
	}

	// -----------------------------------------------------------------------------
	// Phase 2 : Menu interactif semi-automatique
	// -----------------------------------------------------------------------------
	if !promptUserMenu(len(audioFiles)) {
		fmt.Printf("\n%s[INFO] Interruption demandée. À bientôt.%s\n", ANSI_Orange, ANSI_Reset)
		return
	}

	// -----------------------------------------------------------------------------
	// Phase 3 : Traitement séquentiel de la file d'attente
	// -----------------------------------------------------------------------------
	processAudioBatch(audioFiles, inputDir, quarantineDir, whisperCliPath, whisperModelPath, notebookID, semanticRules)

	fmt.Printf("\n%s=== Traitement terminé avec succès. Fermeture du script. ===%s\n", ANSI_Vert, ANSI_Reset)

	// -----------------------------------------------------------------------------
	// Phase 4 : Nettoyage Mémoire stricte
	// -----------------------------------------------------------------------------
	runtime.GC()
	time.Sleep(3 * time.Second)
}

// =============================================================================
// FONCTIONS MÉTIERS ET ORCHESTRATION
// =============================================================================

// Parcours du dossier pour isoler uniquement les formats audio cibles
func scanDirectoryForAudio(inputDir string) []os.DirEntry {
	files, _ := os.ReadDir(inputDir)
	var audioFiles []os.DirEntry

	for _, file := range files {
		if file.IsDir() {
			continue
		}
		ext := strings.ToLower(filepath.Ext(file.Name()))
		if ext == ".mp3" || ext == ".wav" || ext == ".m4a" || ext == ".ogg" {
			audioFiles = append(audioFiles, file)
		}
	}
	return audioFiles
}

// Affichage d'un menu de validation bloquant
func promptUserMenu(fileCount int) bool {
	reader := bufio.NewReader(os.Stdin)
	for {
		fmt.Printf("\n%s=== %d fichier(s) audio détecté(s) ===%s\n", ANSI_Orange, fileCount, ANSI_Reset)
		fmt.Println("1. Lancer la transcription")
		fmt.Println("2. Quitter")
		fmt.Print(ANSI_Vert + "Sélectionnez une option (1 ou 2) : " + ANSI_Reset)

		input, _ := reader.ReadString('\n')
		input = strings.TrimSpace(input)

		if input == "1" {
			return true
		} else if input == "2" {
			return false
		} else {
			fmt.Printf("%s[ERREUR] Saisie invalide. Veuillez taper 1 ou 2.%s\n", ANSI_Rouge, ANSI_Reset)
		}
	}
}

// Orchestration principale : Transcription, Lecture, Rendu Markdown, API et Nettoyage
func processAudioBatch(files []os.DirEntry, inputDir, quarantineDir, cliPath, modelPath, notebookID string, semanticRules map[string]string) {
	regexDate := regexp.MustCompile(`^(\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})`)
	counter := 1

	for _, file := range files {
		audioPath := filepath.Join(inputDir, file.Name())
		fmt.Printf("\n%s[ÉTAPE] Traitement Air-Gapped : %s%s\n", ANSI_Vert, file.Name(), ANSI_Reset)

		txtPath := transcribeAudioSilently(audioPath, inputDir, cliPath, modelPath)
		if txtPath == "" {
			continue
		}

		transcriptionBytes, err := os.ReadFile(txtPath)
		if err != nil {
			fmt.Printf("%s[ERREUR] Impossible de lire la transcription générée.%s\n", ANSI_Rouge, ANSI_Reset)
			continue
		}

		rawText := strings.TrimSpace(strings.ReplaceAll(string(transcriptionBytes), "\r\n", "\n"))
		matches := regexDate.FindStringSubmatch(file.Name())
		var dateStr, timeStr, noteTitle string

		if len(matches) >= 6 {
			dateStr = fmt.Sprintf("%s/%s/20%s", matches[3], matches[2], matches[1])
			timeStr = fmt.Sprintf("%s:%s", matches[4], matches[5])
			noteTitle = fmt.Sprintf("%d - transcrit - %s %s", counter, dateStr, timeStr)
		} else {
			dateStr = "Date Inconnue"
			timeStr = "--:--"
			noteTitle = fmt.Sprintf("%d - transcrit - Date Inconnue", counter)
		}

		formattedMarkdown := generateSemanticMarkdown(noteTitle, dateStr, timeStr, rawText, semanticRules)
		fmt.Printf("%s[INFO] Injection Markdown structurée dans Joplin...%s\n", ANSI_Vert, ANSI_Reset)

		success := sendToJoplin(notebookID, noteTitle, formattedMarkdown)

		// Règle Anti-Destruction stricte (Safe Execution : Move-Only)
		if success {
			moveToQuarantine(txtPath, quarantineDir)
			moveToQuarantine(audioPath, quarantineDir)
			fmt.Printf("%s[OK] Traitement terminé et sécurisé (Move-Only).%s\n", ANSI_Vert, ANSI_Reset)
			counter++
		}
	}
}

// Génération ou chargement du référentiel sémantique externe (Agnosticisme métier)
func loadOrGenerateSemanticConfig(configPath string) map[string]string {
	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		fmt.Printf("%s[INFO] Création du fichier de configuration sémantique (Emojis inclus)...%s\n", ANSI_Orange, ANSI_Reset)
		defaultConfig := SemanticConfig{
			Note: "Fichier de règles sémantiques. Les clés (contenant l'emoji et le thème) seront injectées en tant que sous-titres dans la note Markdown si la Regex correspondante est détectée dans la transcription. Sensibilité à la casse ignorée.",
			Keywords: map[string]string{
				"📊 Comptabilité & Finance":  `(?i)\b(comptabilité|bilan|fiscalité|actif|passif|trésorerie|expert-comptable|TVA|fiscale|liasse)\b`,
				"⚖️ Droit & Juridique":      `(?i)\b(droit|loi|juridique|contrat|législation|décret|jurisprudence|pénal|civil)\b`,
				"💻 IT, Math & Dev":          `(?i)\b(informatique|python|javascript|rust|go|algorithme|serveur|code|base de données|mathématiques|équation|cuda)\b`,
				"♟️ Stratégie & Management": `(?i)\b(stratégie|management|leadership|objectif|kpi|organisation|gouvernance|supply chain|logistique)\b`,
				"💪 Santé & Fitness":         `(?i)\b(santé|fitness|entraînement|nutrition|métabolisme|physiologie|musculation)\b`,
				"🧠 Culture & Dev Perso":     `(?i)\b(culture générale|développement personnel|habitude|discipline|philosophie|apprentissage)\b`,
			},
		}

		file, _ := os.Create(configPath)
		defer file.Close()
		encoder := json.NewEncoder(file)
		encoder.SetIndent("", "  ")
		encoder.Encode(defaultConfig)
		return defaultConfig.Keywords
	}

	file, err := os.Open(configPath)
	if err != nil {
		return make(map[string]string)
	}
	defer file.Close()

	var config SemanticConfig
	json.NewDecoder(file).Decode(&config)
	return config.Keywords
}

// Moteur de rendu texte vers Markdown H1/H3
func generateSemanticMarkdown(title, dateStr, timeStr, rawText string, semanticRules map[string]string) string {
	var detectedDomains []string
	for domain, pattern := range semanticRules {
		matched, _ := regexp.MatchString(pattern, rawText)
		if matched {
			detectedDomains = append(detectedDomains, domain)
		}
	}

	var builder strings.Builder
	builder.WriteString(fmt.Sprintf("# %s\n---\n", title))
	builder.WriteString(fmt.Sprintf("**Date de l'enregistrement :** %s à %s\n\n", dateStr, timeStr))

	if len(detectedDomains) > 0 {
		builder.WriteString("## 🎯 Thématiques identifiées\n\n")
		for _, d := range detectedDomains {
			builder.WriteString(fmt.Sprintf("### %s\n", d))
		}
		builder.WriteString("\n---\n")
	}

	builder.WriteString("## 📝 Transcription\n\n")
	builder.WriteString(rawText)

	return builder.String()
}

// Encapsulation Headless du processus C++ CUDA avec Goroutine bloquante
func transcribeAudioSilently(audioPath, inputDir, cliPath, modelPath string) string {
	cmd := exec.Command(cliPath, "-m", modelPath, "-f", audioPath, "-l", "auto", "-t", "4", "-otxt")
	cmd.Dir = filepath.Dir(cliPath)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}

	if err := cmd.Start(); err != nil {
		return ""
	}

	done := make(chan error, 1)
	go func() {
		done <- cmd.Wait()
	}()

	ticker := time.NewTicker(200 * time.Millisecond)
	defer ticker.Stop()

	spinner := []string{"|", "/", "-", "\\"}
	spinIdx := 0

	fmt.Print(ANSI_Orange + "[INFO] Transcription GPU en cours ")

	for {
		select {
		case err := <-done:
			if err == nil {
				fmt.Printf("\r%s[INFO] Transcription GPU en cours ... Terminée !%s          \n", ANSI_Vert, ANSI_Reset)
			}
			time.Sleep(2 * time.Second) // Flush SSD
			base := strings.TrimSuffix(filepath.Base(audioPath), filepath.Ext(audioPath))
			matches, _ := filepath.Glob(filepath.Join(inputDir, base+"*.txt"))
			if len(matches) > 0 {
				return matches[0]
			}
			return ""

		case <-ticker.C:
			fmt.Printf("\r%s[INFO] Transcription GPU en cours %s %s", ANSI_Orange, spinner[spinIdx], ANSI_Reset)
			spinIdx = (spinIdx + 1) % len(spinner)
		}
	}
}

// =============================================================================
// REQUÊTES HTTP ET SYSTÈME DE FICHIERS
// =============================================================================

func getJoplinFolderID(folderName string) string {
	url := fmt.Sprintf("http://localhost:%s/folders?token=%s", JoplinPort, JoplinToken)
	client := http.Client{Timeout: 3 * time.Second}

	resp, err := client.Get(url)
	if err != nil || resp.StatusCode != 200 {
		return ""
	}
	defer resp.Body.Close()

	var data JoplinFoldersResponse
	json.NewDecoder(resp.Body).Decode(&data)

	for _, folder := range data.Items {
		if folder.Title == folderName {
			return folder.ID
		}
	}
	return ""
}

func sendToJoplin(folderID, title, body string) bool {
	note := JoplinNote{ParentID: folderID, Title: title, Body: body}
	jsonData, _ := json.Marshal(note)

	url := fmt.Sprintf("http://localhost:%s/notes?token=%s", JoplinPort, JoplinToken)
	req, _ := http.NewRequest("POST", url, bytes.NewBuffer(jsonData))
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{Timeout: 5 * time.Second}
	resp, err := client.Do(req)

	if err != nil || resp.StatusCode != 200 {
		return false
	}

	io.Copy(io.Discard, resp.Body)
	defer resp.Body.Close()

	return true
}

func findExecutable(homeDir string) string {
	paths := []string{
		filepath.Join(homeDir, "Documents", "whisper.cpp", "build", "bin", "whisper-cli.exe"),
		filepath.Join(homeDir, "Documents", "whisper.cpp", "build", "bin", "Release", "whisper-cli.exe"),
	}

	for _, p := range paths {
		if _, err := os.Stat(p); err == nil {
			return p
		}
	}
	return ""
}

func findModel(modelsDir string) string {
	files, _ := os.ReadDir(modelsDir)
	for _, f := range files {
		name := strings.ToLower(f.Name())
		if !f.IsDir() && strings.Contains(name, "large-v3-turbo") && (strings.HasSuffix(name, ".bin") || strings.HasSuffix(name, ".gguf")) {
			return filepath.Join(modelsDir, f.Name())
		}
	}
	return ""
}

// Isolement strict des fichiers, interdiction d'utiliser os.Remove
func moveToQuarantine(src, quarantineDir string) {
	filename := filepath.Base(src)
	ts := time.Now().Format("20060102_150405")
	ext := filepath.Ext(filename)
	safeName := fmt.Sprintf("%s_%s%s", strings.TrimSuffix(filename, ext), ts, ext)

	os.Rename(src, filepath.Join(quarantineDir, safeName))
}
