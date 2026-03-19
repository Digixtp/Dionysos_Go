/*
Titre : Orchestrateur de Transcription Hors-Ligne (Buzz & Joplin)
Auteur : Digixtp
Version : 1.0.22
Objectif : Moteur compilé agnostique. Restauration de l'interception stricte des crashs CLI (Fail Fast), retrait des paramètres incompatibles, héritage natif de l'accélération matérielle.

Rappel PowerShell de compilation :
cd "C:\Users\jorda\Documents\Dionysos_Go"
go build -ldflags="-s -w" -o Wispr.Bridge.exe Wispr.Bridge.go
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
	"strings"
	"time"
)

// ==========================================
// Variables Globales et Configuration
// ==========================================
const (
	JoplinAPIToken = "034a2ee109ad401f8246296d7def3edc28dd73accdb49744f1438227784990e56d4e011cdbaa903282931f185dc5ebaee6e9fa5f85de2951612c4f7d0deac651"

	ColorGreen = "\033[92m"
	ColorCyan  = "\033[96m"
	ColorRed   = "\033[91m"
	ColorReset = "\033[0m"
)

type JoplinResponse struct {
	ID    string `json:"id"`
	Title string `json:"title"`
}

type JoplinFolderResponse struct {
	Items []JoplinResponse `json:"items"`
}

type FolderPayload struct {
	Title string `json:"title"`
}

// ==========================================
// Bloc 1 : Utilitaires de Sécurité et Environnement
// ==========================================
func getEnvOrPanic(key string) string {
	value := os.Getenv(key)
	if value == "" {
		fmt.Printf("%s[Erreur Critique] La variable %s est manquante.%s\n", ColorRed, key, ColorReset)
		waitForExit()
		os.Exit(1)
	}
	return value
}

func waitForExit() {
	fmt.Println("\nAppuyez sur Entrée pour quitter...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
}

// ==========================================
// Bloc 2 : Vérification Préventive (Fail Fast)
// ==========================================
func pingJoplinAPI() {
	fmt.Printf("%s[Step] Vérification de la disponibilité de l'API Joplin (Ping)...%s\n", ColorCyan, ColorReset)
	url := fmt.Sprintf("http://localhost:41184/ping?token=%s", JoplinAPIToken)

	resp, err := http.Get(url)
	if err != nil || resp.StatusCode != http.StatusOK {
		fmt.Printf("%s[Erreur Critique] L'API Joplin ne répond pas.%s\n", ColorRed, ColorReset)
		waitForExit()
		os.Exit(1)
	}
	defer resp.Body.Close()
	fmt.Printf("%s[Ok] API Joplin en ligne.%s\n", ColorGreen, ColorReset)
}

func findBuzzExecutable(userProfile string) string {
	pathsToTry := []string{
		filepath.Join(os.Getenv("ProgramFiles"), "Buzz", "Buzz.exe"),
		filepath.Join(os.Getenv("ProgramFiles(x86)"), "Buzz", "Buzz.exe"),
		filepath.Join(userProfile, "AppData", "Local", "Programs", "Buzz", "Buzz.exe"),
	}
	for _, p := range pathsToTry {
		if _, err := os.Stat(p); err == nil {
			return p
		}
	}
	fmt.Printf("%s[Erreur Critique] Buzz.exe introuvable.%s\n", ColorRed, ColorReset)
	waitForExit()
	os.Exit(1)
	return ""
}

// ==========================================
// Bloc 3 : Lecture des Paramètres Natifs (settings.json)
// ==========================================
func getBuzzDefaultExportDir(userProfile string) string {
	settingsPath := filepath.Join(userProfile, "AppData", "Local", "Buzz", "Buzz", "settings.json")
	if data, err := os.ReadFile(settingsPath); err == nil {
		var settings map[string]interface{}
		if err := json.Unmarshal(data, &settings); err == nil {
			if dir, ok := settings["defaultExportDirectory"].(string); ok && dir != "" {
				return dir
			}
		}
	}
	return filepath.Join(userProfile, "Documents", "Buzz")
}

// ==========================================
// Bloc 4 : Nettoyage Chirurgical de la File d'Attente
// ==========================================
func purgerHistoriqueBuzz(userProfile string) {
	dossierQuarantaine := filepath.Join(userProfile, "Desktop", "fichier à supprimer")
	dossierCree := false

	cibles := []string{
		filepath.Join(userProfile, "AppData", "Local", "Buzz", "Buzz", "database.sqlite"),
		filepath.Join(userProfile, "AppData", "Roaming", "Buzz", "database.sqlite"),
	}

	for _, fichier := range cibles {
		if _, err := os.Stat(fichier); err == nil {
			if !dossierCree {
				os.MkdirAll(dossierQuarantaine, os.ModePerm)
				dossierCree = true
			}
			nomSecurise := fmt.Sprintf("database_historique_%s.sqlite", time.Now().Format("150405"))
			os.Rename(fichier, filepath.Join(dossierQuarantaine, nomSecurise))
		}
	}
}

// ==========================================
// Bloc 5 : Pont Joplin Dynamique
// ==========================================
func getOrCreateJoplinFolder() string {
	fmt.Printf("%s[Step] Synchronisation du dossier 'Wispr_Bridge' dans Joplin...%s\n", ColorCyan, ColorReset)
	url := fmt.Sprintf("http://localhost:41184/folders?token=%s", JoplinAPIToken)

	resp, err := http.Get(url)
	if err != nil {
		fmt.Printf("%s[Erreur] Impossible d'interroger l'API Joplin.%s\n", ColorRed, ColorReset)
		waitForExit()
		os.Exit(1)
	}
	defer resp.Body.Close()

	bodyBytes, _ := io.ReadAll(resp.Body)
	var folderData JoplinFolderResponse
	if json.Unmarshal(bodyBytes, &folderData) == nil && len(folderData.Items) > 0 {
		for _, folder := range folderData.Items {
			if folder.Title == "Wispr_Bridge" {
				return folder.ID
			}
		}
	} else {
		var directArray []JoplinResponse
		if json.Unmarshal(bodyBytes, &directArray) == nil {
			for _, folder := range directArray {
				if folder.Title == "Wispr_Bridge" {
					return folder.ID
				}
			}
		}
	}

	payload := FolderPayload{Title: "Wispr_Bridge"}
	jsonData, _ := json.Marshal(payload)
	postResp, _ := http.Post(url, "application/json", bytes.NewBuffer(jsonData))
	defer postResp.Body.Close()

	var newFolder JoplinResponse
	json.NewDecoder(postResp.Body).Decode(&newFolder)
	return newFolder.ID
}

// ==========================================
// Bloc 6 : Classement, Archivage et Application Règle Anti-Destruction
// ==========================================
func archiverTranscriptionBrute(cible string, userProfile string) {
	if _, err := os.Stat(cible); err == nil {
		dossierArchivage := filepath.Join(userProfile, "Desktop", "Fichier des transcriptions brutes")
		os.MkdirAll(dossierArchivage, os.ModePerm)
		nomSecurise := fmt.Sprintf("%s_%s", time.Now().Format("20060102_150405"), filepath.Base(cible))
		os.Rename(cible, filepath.Join(dossierArchivage, nomSecurise))
	}
}

func classerAudioTraite(cheminAudio string, dossierParent string) {
	nomDossierCible := fmt.Sprintf("Audio traités_%s", time.Now().Format("02.01.2006"))
	dossierCible := filepath.Join(dossierParent, nomDossierCible)
	os.MkdirAll(dossierCible, os.ModePerm)
	os.Rename(cheminAudio, filepath.Join(dossierCible, filepath.Base(cheminAudio)))
	fmt.Printf("%s[Ok] Audio source classé dans '%s'.%s\n", ColorGreen, nomDossierCible, ColorReset)
}

func purgerFichiersResiduels(repertoires []string, nomBase string, userProfile string) {
	dossierQuarantaine := filepath.Join(userProfile, "Desktop", "fichier à supprimer")
	dossierCree := false
	extensionsIndesirables := []string{".srt", ".vtt", ".json", ".csv"}

	for _, dossier := range repertoires {
		for _, ext := range extensionsIndesirables {
			matches, _ := filepath.Glob(filepath.Join(dossier, nomBase+"*"+ext))
			for _, match := range matches {
				if !dossierCree {
					os.MkdirAll(dossierQuarantaine, os.ModePerm)
					dossierCree = true
				}
				nomSecurise := fmt.Sprintf("residu_%s_%s", time.Now().Format("150405"), filepath.Base(match))
				os.Rename(match, filepath.Join(dossierQuarantaine, nomSecurise))
			}
		}
	}
}

// ==========================================
// Bloc 7 : Formatage Markdown Professionnel
// ==========================================
func cleanAndFormatMarkdown(rawText string, originalFileName string) string {
	re := regexp.MustCompile(`\[\d{2}:\d{2}:\d{2}\.\d{3}\]\s*|\[\d{2}:\d{2}\.\d{3}\]\s*`)
	cleanedText := strings.TrimSpace(re.ReplaceAllString(rawText, ""))

	markdownStruct := fmt.Sprintf("# Transcription de %s\n\n", originalFileName)
	markdownStruct += fmt.Sprintf("**Date du traitement :** %s\n", time.Now().Format("02/01/2006 à 15h04"))
	markdownStruct += "**Tags :** #transcription #buzz\n\n---\n\n"
	markdownStruct += cleanedText
	return markdownStruct
}

// ==========================================
// Bloc 8 : Sécurité Réseau (Air-Gap)
// ==========================================
func enforceOfflineMode(buzzExePath string) {
	psScript := fmt.Sprintf(`$ruleName = "Block_Buzz_Offline"; if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName $ruleName -Direction Outbound -Program "%s" -Action Block | Out-Null }`, buzzExePath)
	exec.Command("powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript).Run()
}

// ==========================================
// Bloc 9 : Génération du Nom Cible
// ==========================================
func generateFilename(audioFile string, baseDir string) string {
	info, err := os.Stat(audioFile)
	if err != nil {
		return "X - transcrit - erreur_date"
	}

	dateStr := info.ModTime().Format("02-01-2006_15h04")
	files, _ := os.ReadDir(baseDir)
	count := 1
	for _, f := range files {
		if !f.IsDir() && strings.Contains(f.Name(), "- transcrit -") {
			count++
		}
	}
	return fmt.Sprintf("%d - transcrit - %s", count, dateStr)
}

// ==========================================
// Bloc 10 : Export vers l'API Joplin
// ==========================================
func exportToJoplin(title string, body string, folderID string) {
	url := fmt.Sprintf("http://localhost:41184/notes?token=%s", JoplinAPIToken)
	payload := map[string]string{"parent_id": folderID, "title": title, "body": body}
	jsonData, _ := json.Marshal(payload)

	req, _ := http.NewRequest("POST", url, bytes.NewBuffer(jsonData))
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{}
	resp, err := client.Do(req)

	if err != nil || (resp.StatusCode != http.StatusOK && resp.StatusCode != http.StatusCreated) {
		fmt.Printf("\n%s[Erreur] L'API a refusé l'injection.%s\n", ColorRed, ColorReset)
		if resp != nil {
			resp.Body.Close()
		}
		return
	}
	defer resp.Body.Close()

	var joplinResp JoplinResponse
	json.NewDecoder(resp.Body).Decode(&joplinResp)
	fmt.Printf("\n%s[Ok] Note générée dans Joplin. ID d'Audit : %s%s\n", ColorGreen, joplinResp.ID, ColorReset)
}

// ==========================================
// Point d'Entrée Principal (Moteur d'Exécution)
// ==========================================
func main() {
	userProfile := getEnvOrPanic("USERPROFILE")
	baseDir := filepath.Join(userProfile, "Documents", "Wispr_Bridge")
	os.MkdirAll(baseDir, os.ModePerm)

	fmt.Printf("%s=== Moteur Wispr.Bridge v1.0.22 ===%s\n", ColorCyan, ColorReset)

	exec.Command("taskkill", "/IM", "Buzz.exe", "/F", "/T").Run()

	pingJoplinAPI()
	buzzExePath := findBuzzExecutable(userProfile)
	enforceOfflineMode(buzzExePath)
	dynamicFolderID := getOrCreateJoplinFolder()
	buzzExportDir := getBuzzDefaultExportDir(userProfile)

	fmt.Print("\nVeuillez coller le chemin du dossier contenant les audios : ")
	reader := bufio.NewReader(os.Stdin)
	inputPath, _ := reader.ReadString('\n')
	inputPath = strings.Trim(strings.TrimSpace(inputPath), "\"'")

	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		fmt.Printf("%s[Erreur] Dossier introuvable.%s\n", ColorRed, ColorReset)
		waitForExit()
		os.Exit(1)
	}

	entries, err := os.ReadDir(inputPath)
	if err != nil {
		fmt.Printf("%s[Erreur] Lecture du dossier impossible.%s\n", ColorRed, ColorReset)
		waitForExit()
		os.Exit(1)
	}

	supportedExtensions := map[string]bool{".mp3": true, ".wav": true, ".m4a": true, ".mp4": true, ".flac": true, ".wma": true}
	fichiersTraites := 0

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		ext := strings.ToLower(filepath.Ext(entry.Name()))
		if !supportedExtensions[ext] {
			continue
		}

		audioFile := filepath.Join(inputPath, entry.Name())
		filename := generateFilename(audioFile, baseDir)
		nomBaseGenere := strings.TrimSuffix(entry.Name(), ext)

		fmt.Printf("\n%s[Step] Traitement de : %s%s\n", ColorCyan, entry.Name(), ColorReset)

		exec.Command("taskkill", "/IM", "Buzz.exe", "/F", "/T").Run()
		time.Sleep(1 * time.Second)
		purgerHistoriqueBuzz(userProfile)

		// Exécution épurée : héritage natif des paramètres matériels de l'interface graphique Buzz
		cmd := exec.Command(buzzExePath, "add", "--task", "transcribe", "--model-size", "large-v3-turbo", audioFile)

		var errBuf bytes.Buffer
		cmd.Stdout = os.Stdout
		cmd.Stderr = io.MultiWriter(os.Stderr, &errBuf)

		// Rétablissement du Fail Fast sur la commande CLI
		if err := cmd.Run(); err != nil {
			fmt.Printf("\n%s[Erreur] Commande rejetée par Buzz : %v%s\n", ColorRed, err, ColorReset)
			fmt.Printf("%s[Détail technique] %s%s\n", ColorRed, strings.TrimSpace(errBuf.String()), ColorReset)
			continue
		}

		dossiersAChercher := []string{buzzExportDir, inputPath, baseDir}
		frames := []string{"[■□□□□□]", "[■■□□□□]", "[■■■□□□]", "[■■■■□□]", "[■■■■■□]", "[■■■■■■]"}
		fichierTrouve := false
		cheminFinalTxt := ""
		start := time.Now()

		for tentative := 0; tentative < 3600; tentative++ {
			elapsed := time.Since(start).Round(time.Second)
			fmt.Printf("\r\033[K%s[Info] Transcription IA en cours %s (Temps écoulé: %v)%s", ColorCyan, frames[tentative%len(frames)], elapsed, ColorReset)

			for _, dossier := range dossiersAChercher {
				matches, _ := filepath.Glob(filepath.Join(dossier, nomBaseGenere+"*.txt"))
				if len(matches) > 0 {
					cheminFinalTxt = matches[0]
					fichierTrouve = true
					break
				}
			}

			if fichierTrouve {
				time.Sleep(1 * time.Second)
				break
			}
			time.Sleep(1 * time.Second)
		}

		if !fichierTrouve {
			fmt.Printf("\n%s[Erreur] Le fichier texte n'a pas été trouvé.%s\n", ColorRed, ColorReset)
			continue
		}

		rawBytes, err := os.ReadFile(cheminFinalTxt)
		if err != nil {
			fmt.Printf("\n%s[Erreur] Lecture impossible : %v%s\n", ColorRed, err, ColorReset)
			continue
		}

		formattedMarkdown := cleanAndFormatMarkdown(string(rawBytes), entry.Name())
		exportToJoplin(filename, formattedMarkdown, dynamicFolderID)

		fichiersTraites++
		archiverTranscriptionBrute(cheminFinalTxt, userProfile)
		purgerFichiersResiduels(dossiersAChercher, nomBaseGenere, userProfile)
		classerAudioTraite(audioFile, inputPath)
	}

	fmt.Printf("\n%s[Step] Clôture finale des processus...%s\n", ColorCyan, ColorReset)
	exec.Command("taskkill", "/IM", "Buzz.exe", "/F", "/T").Run()

	fmt.Printf("%s=== Terminé : %d fichier(s) traité(s) ===%s\n", ColorGreen, fichiersTraites, ColorReset)
	waitForExit()
}
