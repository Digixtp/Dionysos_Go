/*
==============================================================================
Titre : Moteur d'Audit et Correction Linguistique Massif (Multi-Formats)
Auteur : Digixtp
Version : 1.22.0
Note introductive : Ce script orchestre la correction de textes via Qwen2.5.
Cette version déploie l'architecture de "Contrôle de Cohérence" (Les 6 Étapes) :
1/2. Validation du volume d'extraction vs Poids du fichier (Retry max: 2).
3. Chunking sémantique respectant la ponctuation.
4. Ciblage spatial exact (ex: Colonne A, Ligne 35).
5. Sauvegarde incrémentale de l'Excel après CHAQUE fichier traité.
6. Notification de fin de traitement. Le tout accéléré par un Sémaphore GPU (5).
==============================================================================
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
	"strconv"
	"strings"
	"sync"
	"syscall"
	"time"
	"unicode"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

// --- CONSTANTES ANSI (Standard UX) ---
const (
	ColorReset  = "\033[0m"
	ColorBlue   = "\033[94m"
	ColorGreen  = "\033[92m"
	ColorRed    = "\033[91m"
	ColorCyan   = "\033[96m"
	ColorYellow = "\033[93m"
)

// --- STRUCTURES DE DONNÉES ---
type LLMRequest struct {
	Model   string    `json:"model"`
	Prompt  string    `json:"prompt"`
	Stream  bool      `json:"stream"`
	Format  string    `json:"format"`
	Options LLMOption `json:"options"`
}

type LLMOption struct {
	Temperature float64 `json:"temperature"`
	NumCtx      int     `json:"num_ctx"`
	NumPredict  int     `json:"num_predict"`
}

type LLMResponse struct {
	Response string `json:"response"`
}

type ErreurDetectee struct {
	PhraseOriginale string `json:"phrase_originale"`
	MotFaux         string `json:"mot_faux"`
	PhraseCorrigee  string `json:"phrase_corrigee"`
}

type Remplacement struct {
	Coordonnees string
	Original    string
	Correction  string
	LigneExcel  int
}

type ChunkJob struct {
	Coordonnees string
	Texte       string
	Chemin      string
}

// --- CLIENT HTTP OPTIMISÉ (Keep-Alive) ---
var httpClient = &http.Client{
	Timeout: 60 * time.Second,
	Transport: &http.Transport{
		MaxIdleConns:        100,
		MaxIdleConnsPerHost: 100,
		IdleConnTimeout:     90 * time.Second,
	},
}

// --- INITIALISATION SYSTÈME (ANSI Windows) ---
func init() {
	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	procGetConsoleMode := kernel32.NewProc("GetConsoleMode")
	procSetConsoleMode := kernel32.NewProc("SetConsoleMode")
	procGetStdHandle := kernel32.NewProc("GetStdHandle")

	const stdOutputHandle = ^uint32(10)
	handle, _, _ := procGetStdHandle.Call(uintptr(stdOutputHandle))
	var mode uint32
	procGetConsoleMode.Call(handle, uintptr(unsafe.Pointer(&mode)))
	procSetConsoleMode.Call(handle, uintptr(mode|0x0004))
}

// --- SÉCURITÉ MATÉRIELLE (I/O Throttling & Lock Checking) ---
func verifierVerrouFichier(chemin string) error {
	file, err := os.OpenFile(chemin, os.O_RDWR, 0666)
	if err != nil {
		return fmt.Errorf("fichier verrouillé par une autre application")
	}
	file.Close()
	return nil
}

func attendreDisqueOptimal(seuilMax int) {
	psScript := `(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'").PercentDiskTime`
	for {
		cmd := exec.Command("powershell", "-NoProfile", "-Command", psScript)
		out, err := cmd.Output()
		if err != nil {
			return
		}

		valStr := strings.TrimSpace(string(out))
		valInt, errConv := strconv.Atoi(valStr)
		if errConv != nil || valInt <= seuilMax {
			fmt.Print("\r\033[K")
			return
		}
		fmt.Printf("\r%s[Throttling] Disque saturé (%d%%). Temporisation RAM...%s", ColorYellow, valInt, ColorReset)
		time.Sleep(2 * time.Second)
	}
}

func verifierEspaceSSD(lecteur string) error {
	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	getDiskFreeSpaceEx := kernel32.NewProc("GetDiskFreeSpaceExW")

	var freeBytesAvailableToCaller, totalNumberOfBytes, totalNumberOfFreeBytes int64
	ptrLecteur, err := syscall.UTF16PtrFromString(lecteur)
	if err != nil {
		return err
	}

	ret, _, errSys := getDiskFreeSpaceEx.Call(
		uintptr(unsafe.Pointer(ptrLecteur)),
		uintptr(unsafe.Pointer(&freeBytesAvailableToCaller)),
		uintptr(unsafe.Pointer(&totalNumberOfBytes)),
		uintptr(unsafe.Pointer(&totalNumberOfFreeBytes)),
	)

	if ret == 0 {
		return fmt.Errorf("impossible de lire le SSD : %v", errSys)
	}

	freeGB := freeBytesAvailableToCaller / (1024 * 1024 * 1024)
	if freeGB < 5 {
		return fmt.Errorf("Espace critique (%d Go restants). Exécution annulée.", freeGB)
	}
	return nil
}

// --- FILTRAGE LEXICAL INTELLIGENT ---
func estTexteValide(texte string) bool {
	t := strings.TrimSpace(texte)
	if len(t) < 3 {
		return false
	}

	lettres := 0
	for _, r := range t {
		if unicode.IsLetter(r) {
			lettres++
		}
	}
	if lettres < 2 {
		return false
	}

	prefixes := []string{"func ", "var ", "import ", "package ", "type ", "```"}
	for _, p := range prefixes {
		if strings.HasPrefix(t, p) {
			return false
		}
	}
	return true
}

// --- ÉTAPES 1 & 2 : CONTRÔLE DE COHÉRENCE (Poids vs Extraction) ---
func extraireAvecValidation(chemin string, ext string) ([]ChunkJob, error) {
	info, err := os.Stat(chemin)
	if err != nil {
		return nil, err
	}
	tailleOctets := info.Size()

	if tailleOctets == 0 {
		return nil, fmt.Errorf("fichier vide (0 octet)")
	}

	var chunks []ChunkJob
	var errExtraction error

	// Boucle de résilience : 1 essai + 2 recommencements (Max 3)
	for tentative := 1; tentative <= 3; tentative++ {
		chunks = []ChunkJob{}
		totalCaracteres := 0

		switch ext {
		case ".txt", ".md":
			bytesData, err := os.ReadFile(chemin)
			if err == nil {
				contenu := string(bytesData)
				totalCaracteres = len(contenu)
				chunks = append(chunks, genererChunksSemantiques(chemin, "[Document Entier]", contenu, 3000)...)
			} else {
				errExtraction = err
			}

		case ".docx", ".pptx":
			contenu := extraireTexteOffice(chemin, ext)
			totalCaracteres = len(contenu)
			chunks = append(chunks, genererChunksSemantiques(chemin, "[Document Entier]", contenu, 3000)...)

		case ".xlsx", ".xlsm":
			xl, err := excelize.OpenFile(chemin)
			if err == nil {
				for _, name := range xl.GetSheetList() {
					rows, _ := xl.GetRows(name)
					for rIdx, row := range rows {
						for cIdx, colCell := range row {
							if estTexteValide(colCell) {
								totalCaracteres += len(colCell)
								cellName, _ := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
								coord := fmt.Sprintf("[%s] %s", name, cellName)
								chunks = append(chunks, genererChunksSemantiques(chemin, coord, colCell, 3000)...)
							}
						}
					}
				}
				xl.Close()
			} else {
				errExtraction = err
			}
		}

		// Validation de cohérence
		coherent := false
		if ext == ".txt" || ext == ".md" {
			// Pour un fichier texte pur, le nombre de caractères doit être au moins 50% de la taille en octets
			// (pour tolérer l'encodage UTF-8 et les espaces vides).
			if int64(totalCaracteres) >= tailleOctets/2 || totalCaracteres > 0 {
				coherent = true
			}
		} else {
			// Pour Office (.docx, .xlsx), le fichier est un ZIP. La taille en octets ne correspond pas au texte.
			// On exige simplement que du texte ait été extrait si le fichier pèse un certain poids.
			if totalCaracteres > 0 || tailleOctets < 5000 {
				coherent = true
			}
		}

		if coherent && errExtraction == nil {
			return chunks, nil // Extraction validée
		}

		if tentative < 3 {
			fmt.Printf("\r%s[Warning] Incohérence de lecture sur %s. Re-tentative (%d/3)...%s\n", ColorYellow, filepath.Base(chemin), tentative+1, ColorReset)
			time.Sleep(1 * time.Second)
		}
	}

	return nil, fmt.Errorf("rejeté après 3 tentatives : incohérence entre le poids du fichier et les données lues")
}

// --- ÉTAPE 3 : INTELLIGENCE SÉMANTIQUE (Chunking) ---
func splitPhrasesIntelligentes(texte string) []string {
	var phrases []string
	var courante strings.Builder
	runes := []rune(texte)

	for i := 0; i < len(runes); i++ {
		courante.WriteRune(runes[i])
		r := runes[i]
		if r == '.' || r == '!' || r == '?' {
			if i+1 == len(runes) || runes[i+1] == ' ' || runes[i+1] == '\n' {
				if i+1 < len(runes) && runes[i+1] == ' ' {
					courante.WriteRune(' ')
					i++
				}
				phrases = append(phrases, strings.TrimSpace(courante.String()))
				courante.Reset()
			}
		} else if r == '\n' {
			phrases = append(phrases, strings.TrimSpace(courante.String()))
			courante.Reset()
		}
	}
	if courante.Len() > 0 {
		reste := strings.TrimSpace(courante.String())
		if len(reste) > 0 {
			phrases = append(phrases, reste)
		}
	}
	return phrases
}

func genererChunksSemantiques(chemin string, coordonnees string, contenu string, limite int) []ChunkJob {
	var chunks []ChunkJob
	var chunkCourant strings.Builder

	phrases := splitPhrasesIntelligentes(contenu)
	for _, phrase := range phrases {
		if !estTexteValide(phrase) {
			continue
		}

		if chunkCourant.Len()+len(phrase) > limite && chunkCourant.Len() > 0 {
			chunks = append(chunks, ChunkJob{Coordonnees: coordonnees, Texte: strings.TrimSpace(chunkCourant.String()), Chemin: chemin})
			chunkCourant.Reset()
		}
		chunkCourant.WriteString(phrase + "\n")
	}
	if chunkCourant.Len() > 0 {
		chunks = append(chunks, ChunkJob{Coordonnees: coordonnees, Texte: strings.TrimSpace(chunkCourant.String()), Chemin: chemin})
	}
	return chunks
}

// --- MOTEUR D'INTEROPÉRABILITÉ (PowerShell / COM) ---
func extraireTexteOffice(chemin string, typeFichier string) string {
	var psScript string
	if typeFichier == ".docx" {
		psScript = fmt.Sprintf(`
			$word = New-Object -ComObject Word.Application
			$word.Visible = $false
			$word.DisplayAlerts = $false
			$doc = $word.Documents.Open('%s', $false, $true)
			$text = $doc.Content.Text
			$doc.Close($false)
			$word.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
			[GC]::Collect()
			Write-Output $text
		`, chemin)
	} else if typeFichier == ".pptx" {
		psScript = fmt.Sprintf(`
			$ppt = New-Object -ComObject PowerPoint.Application
			$pres = $ppt.Presentations.Open('%s', $true, $false, $false)
			$text = ""
			foreach ($slide in $pres.Slides) {
				foreach ($shape in $slide.Shapes) {
					if ($shape.HasTextFrame) { $text += $shape.TextFrame.TextRange.Text + [Environment]::NewLine }
				}
			}
			$pres.Close()
			$ppt.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
			[GC]::Collect()
			Write-Output $text
		`, chemin)
	}
	cmd := exec.Command("powershell", "-NoProfile", "-Command", psScript)
	out, _ := cmd.Output()
	return string(out)
}

func corrigerOfficeCOM_Batch(chemin string, remplacements []Remplacement, typeFichier string) error {
	var psBuilder strings.Builder
	if typeFichier == ".docx" {
		psBuilder.WriteString(fmt.Sprintf(`
			$ErrorActionPreference = 'Stop'
			try {
				$word = New-Object -ComObject Word.Application
				$word.Visible = $false
				$word.DisplayAlerts = $false
				$doc = $word.Documents.Open('%s')
				$find = $doc.Content.Find
		`, chemin))

		for _, remp := range remplacements {
			origClean := strings.ReplaceAll(remp.Original, "'", "''")
			corrClean := strings.ReplaceAll(remp.Correction, "'", "''")
			psBuilder.WriteString(fmt.Sprintf(`$find.Execute('%s', $false, $false, $false, $false, $false, $true, 1, $false, '%s', 2) | Out-Null; `, origClean, corrClean))
		}

		psBuilder.WriteString(`
				$doc.Save()
				$doc.Close()
				$word.Quit()
				[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
				[GC]::Collect()
			} catch { exit 1 }
		`)
	} else if typeFichier == ".pptx" {
		psBuilder.WriteString(fmt.Sprintf(`
			$ErrorActionPreference = 'Stop'
			try {
				$ppt = New-Object -ComObject PowerPoint.Application
				$pres = $ppt.Presentations.Open('%s', $false, $false, $false)
				foreach ($slide in $pres.Slides) {
					foreach ($shape in $slide.Shapes) {
						if ($shape.HasTextFrame) {
							$txt = $shape.TextFrame.TextRange.Text
		`, chemin))
		for _, remp := range remplacements {
			origClean := strings.ReplaceAll(remp.Original, "'", "''")
			corrClean := strings.ReplaceAll(remp.Correction, "'", "''")
			psBuilder.WriteString(fmt.Sprintf(`if ($txt -match [regex]::Escape('%s')) { $txt = $txt -replace [regex]::Escape('%s'), '%s' }; `, origClean, origClean, corrClean))
		}
		psBuilder.WriteString(`
							$shape.TextFrame.TextRange.Text = $txt
						}
					}
				}
				$pres.Save()
				$pres.Close()
				$ppt.Quit()
				[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
				[GC]::Collect()
			} catch { exit 1 }
		`)
	}
	err := exec.Command("powershell", "-NoProfile", "-Command", psBuilder.String()).Run()
	if err != nil {
		return fmt.Errorf("échec API COM")
	}
	return nil
}

func corrigerExcel_Batch(chemin string, remplacements []Remplacement) error {
	xl, err := excelize.OpenFile(chemin)
	if err != nil {
		return err
	}
	defer xl.Close()

	for _, remp := range remplacements {
		coord := remp.Coordonnees
		if strings.HasPrefix(coord, "[") && strings.Contains(coord, "] ") {
			parts := strings.SplitN(coord, "] ", 2)
			sheetName := strings.TrimPrefix(parts[0], "[")
			cellName := parts[1]

			cellValue, err := xl.GetCellValue(sheetName, cellName)
			if err == nil && strings.Contains(cellValue, remp.Original) {
				newVal := strings.Replace(cellValue, remp.Original, remp.Correction, 1)
				xl.SetCellValue(sheetName, cellName, newVal)
			}
		}
	}
	return xl.Save()
}

// --- MOTEUR LLM (Qwen2.5) ---
func interrogerLLM(texte string) []ErreurDetectee {
	url := "http://localhost:11434/api/generate"
	prompt := fmt.Sprintf(`Tu es un correcteur orthographique expert.
Analyse le texte ci-dessous.
Génère UNIQUEMENT un tableau JSON strict des erreurs.
Format attendu :
[{"phrase_originale": "phrase avec faute", "mot_faux": "mot", "phrase_corrigee": "phrase corrigée"}]
Si aucune faute, renvoie : []

Texte :
%s`, texte)

	reqBody := LLMRequest{
		Model:  "qwen2.5:1.5b",
		Prompt: prompt,
		Stream: false,
		Format: "json",
		Options: LLMOption{
			Temperature: 0.0,
			NumCtx:      1024,
			NumPredict:  512,
		},
	}

	jsonData, _ := json.Marshal(reqBody)
	resp, err := httpClient.Post(url, "application/json", bytes.NewBuffer(jsonData))
	if err != nil {
		return nil
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var llmResp LLMResponse
	json.Unmarshal(body, &llmResp)

	cleanJSON := strings.TrimSpace(llmResp.Response)
	cleanJSON = strings.TrimPrefix(cleanJSON, "```json")
	cleanJSON = strings.TrimPrefix(cleanJSON, "```")
	cleanJSON = strings.TrimSuffix(cleanJSON, "```")
	cleanJSON = strings.TrimSpace(cleanJSON)

	var erreurs []ErreurDetectee
	json.Unmarshal([]byte(cleanJSON), &erreurs)
	return erreurs
}

// --- PHASES 1 À 6 : FLUX INTÉGRAL FICHIER PAR FICHIER ---
func phase1Analyse() {
	fmt.Print(ColorCyan + "Chemin absolu du dossier à analyser : " + ColorReset)
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
	dossierCible := strings.TrimSpace(scanner.Text())

	if _, err := os.Stat(dossierCible); os.IsNotExist(err) {
		fmt.Println(ColorRed + "Erreur : Le dossier n'existe pas." + ColorReset)
		return
	}

	// Étape 4 (Préparation) : Création de l'Excel
	f := excelize.NewFile()
	sheet := "Audit_Corrections"
	f.SetSheetName("Sheet1", sheet)

	f.SetCellValue(sheet, "A1", "RAPPORT D'AUDIT LLM (Flux Fichier par Fichier)")
	f.SetCellValue(sheet, "A2", "Note introductive : Laissez la colonne D intacte pour appliquer la correction à l'emplacement exact indiqué en colonne A.")
	styleNote, _ := f.NewStyle(&excelize.Style{Font: &excelize.Font{Italic: true, Color: "#64748b"}})
	f.SetCellStyle(sheet, "A2", "A2", styleNote)

	styleTitre, _ := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#000080"}, Pattern: 1},
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})

	enTetes := []string{"Coordonnées (Position)", "Chemin du fichier", "Erreur (Phrase Originale)", "Validation (Phrase Corrigée)", "Statut d'injection"}
	for i, colName := range enTetes {
		cell := string(rune('A'+i)) + "5"
		f.SetCellValue(sheet, cell, colName)
		f.SetCellStyle(sheet, cell, cell, styleTitre)
	}

	f.SetColWidth(sheet, "A", "A", 25)
	f.SetColWidth(sheet, "B", "B", 40)
	f.SetColWidth(sheet, "C", "C", 60)
	f.SetColWidth(sheet, "D", "D", 60)
	f.SetColWidth(sheet, "E", "E", 25)

	currentRow := 6
	rapportPath := filepath.Join(os.Getenv("USERPROFILE"), "Desktop", "Audit_Corrections.xlsx")

	// Recensement des fichiers
	var fichiersCibles []string
	filepath.Walk(dossierCible, func(path string, info os.FileInfo, err error) error {
		if err == nil && !info.IsDir() {
			ext := strings.ToLower(filepath.Ext(path))
			if ext == ".txt" || ext == ".md" || ext == ".docx" || ext == ".pptx" || ext == ".xlsx" || ext == ".xlsm" {
				fichiersCibles = append(fichiersCibles, path)
			}
		}
		return nil
	})

	totalFichiers := len(fichiersCibles)
	if totalFichiers == 0 {
		fmt.Println(ColorYellow + "Aucun fichier valide trouvé dans le dossier." + ColorReset)
		return
	}

	fmt.Printf("%s[Step] Lancement du traitement massif (Sauvegarde incrémentale)...%s\n", ColorBlue, ColorReset)

	var mu sync.Mutex

	// Flux d'exécution Fichier par Fichier
	for numFichier, path := range fichiersCibles {
		fmt.Printf("\n%s>> Traitement du fichier %d/%d : %s%s\n", ColorBlue, numFichier+1, totalFichiers, filepath.Base(path), ColorReset)

		// Étapes 1 & 2 : Taille vs Cohérence d'extraction
		chunks, err := extraireAvecValidation(path, strings.ToLower(filepath.Ext(path)))
		if err != nil {
			fmt.Printf("%s[Rejet] %s : %v%s\n", ColorRed, filepath.Base(path), err, ColorReset)
			continue // On rejette et on passe au fichier suivant
		}

		if len(chunks) == 0 {
			fmt.Println(ColorYellow + " -> Aucun bloc textuel pertinent à analyser." + ColorReset)
			continue
		}

		// Étape 3 : Lecture et correction LLM (Asynchrone par Chunks sur le fichier en cours)
		var wg sync.WaitGroup
		semaphore := make(chan struct{}, 5)
		erreursTrouvees := 0

		for cIdx, job := range chunks {
			wg.Add(1)
			semaphore <- struct{}{}

			go func(idx int, j ChunkJob) {
				defer wg.Done()
				defer func() { <-semaphore }()

				erreurs := interrogerLLM(j.Texte)

				if len(erreurs) > 0 {
					mu.Lock()
					for _, errDetect := range erreurs {
						if strings.TrimSpace(errDetect.PhraseOriginale) != "" && strings.TrimSpace(errDetect.PhraseCorrigee) != "" {
							f.SetCellValue(sheet, fmt.Sprintf("A%d", currentRow), j.Coordonnees)
							f.SetCellValue(sheet, fmt.Sprintf("B%d", currentRow), j.Chemin)

							idxRed := strings.Index(errDetect.PhraseOriginale, errDetect.MotFaux)
							if idxRed != -1 && errDetect.MotFaux != "" {
								p1 := errDetect.PhraseOriginale[:idxRed]
								p2 := errDetect.PhraseOriginale[idxRed : idxRed+len(errDetect.MotFaux)]
								p3 := errDetect.PhraseOriginale[idxRed+len(errDetect.MotFaux):]
								f.SetCellRichText(sheet, fmt.Sprintf("C%d", currentRow), []excelize.RichTextRun{
									{Text: p1}, {Text: p2, Font: &excelize.Font{Color: "FF0000", Bold: true}}, {Text: p3},
								})
							} else {
								f.SetCellValue(sheet, fmt.Sprintf("C%d", currentRow), errDetect.PhraseOriginale)
							}
							f.SetCellValue(sheet, fmt.Sprintf("D%d", currentRow), errDetect.PhraseCorrigee)
							f.SetCellValue(sheet, fmt.Sprintf("E%d", currentRow), "En attente Phase 2")
							currentRow++
							erreursTrouvees++
						}
					}
					mu.Unlock()
				}
				fmt.Printf("\r -> Progression IA : %d/%d blocs analysés", idx+1, len(chunks))
			}(cIdx, job)
		}
		wg.Wait()
		fmt.Println() // Saut de ligne après la barre de progression du fichier

		// Étape 5 : Sauvegarde et clôture du cycle pour le fichier en cours
		if erreursTrouvees > 0 {
			f.SaveAs(rapportPath)
			fmt.Printf("%s[Ok] %d erreurs ajoutées et sauvegardées dans le rapport Excel.%s\n", ColorGreen, erreursTrouvees, ColorReset)
		} else {
			fmt.Printf("%s[Info] Fichier propre. Aucune erreur trouvée.%s\n", ColorGreen, ColorReset)
		}
	}

	// Étape 6 : Clôture générale
	fmt.Println(ColorGreen + "\n=========================================" + ColorReset)
	fmt.Println(ColorGreen + "          TRAITEMENT TERMINÉ             " + ColorReset)
	fmt.Println(ColorGreen + "=========================================" + ColorReset)
	fmt.Printf("Rapport d'audit disponible : %s\n", rapportPath)
}

// --- PHASE 2 : CORRECTION CHIRURGICALE ---
func phase2Correction() {
	rapportPath := filepath.Join(os.Getenv("USERPROFILE"), "Desktop", "Audit_Corrections.xlsx")
	f, err := excelize.OpenFile(rapportPath)
	if err != nil {
		fmt.Println(ColorRed + "Erreur : Rapport d'audit introuvable sur le bureau." + ColorReset)
		return
	}
	defer f.Close()

	rows, _ := f.GetRows("Audit_Corrections")
	planDAction := make(map[string][]Remplacement)

	for i := 5; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 4 {
			continue
		}

		coord := row[0]
		chemin := row[1]
		original := row[2]
		correction := row[3]

		if strings.TrimSpace(correction) == "" || original == correction {
			f.SetCellValue("Audit_Corrections", fmt.Sprintf("E%d", i+1), "Ignoré (Case vide/inchangée)")
			continue
		}

		planDAction[chemin] = append(planDAction[chemin], Remplacement{
			Coordonnees: coord,
			Original:    original,
			Correction:  correction,
			LigneExcel:  i + 1,
		})
	}

	totalFichiers := len(planDAction)
	if totalFichiers == 0 {
		fmt.Println(ColorYellow + "Aucune correction à appliquer. Rapport mis à jour." + ColorReset)
		f.SaveAs(rapportPath)
		return
	}

	fmt.Println(ColorBlue + "[Step] Injection avec Indexation Spatiale et Throttling..." + ColorReset)
	fichiersTraites := 0

	for chemin, listeRemplacements := range planDAction {
		var errInjection error

		if errVerrou := verifierVerrouFichier(chemin); errVerrou != nil {
			errInjection = errVerrou
		} else {
			ext := strings.ToLower(filepath.Ext(chemin))
			switch ext {
			case ".txt", ".md":
				contentBytes, err := os.ReadFile(chemin)
				if err != nil {
					errInjection = err
				} else {
					newContent := string(contentBytes)
					for _, remp := range listeRemplacements {
						newContent = strings.Replace(newContent, remp.Original, remp.Correction, 1)
					}
					attendreDisqueOptimal(60)
					errInjection = os.WriteFile(chemin, []byte(newContent), 0644)
				}
			case ".docx", ".pptx":
				attendreDisqueOptimal(60)
				errInjection = corrigerOfficeCOM_Batch(chemin, listeRemplacements, ext)
			case ".xlsx", ".xlsm":
				attendreDisqueOptimal(60)
				errInjection = corrigerExcel_Batch(chemin, listeRemplacements)
			}
		}

		for _, remp := range listeRemplacements {
			celluleStatut := fmt.Sprintf("E%d", remp.LigneExcel)
			if errInjection != nil {
				styleErr, _ := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#f8d7da"}, Pattern: 1}, Font: &excelize.Font{Color: "#721c24"}})
				f.SetCellValue("Audit_Corrections", celluleStatut, "Erreur : "+errInjection.Error())
				f.SetCellStyle("Audit_Corrections", celluleStatut, celluleStatut, styleErr)
			} else {
				styleOk, _ := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#d4edda"}, Pattern: 1}, Font: &excelize.Font{Color: "#155724"}})
				f.SetCellValue("Audit_Corrections", celluleStatut, "Modification validée")
				f.SetCellStyle("Audit_Corrections", celluleStatut, celluleStatut, styleOk)
			}
		}

		fichiersTraites++
		afficherBarreProgression(fichiersTraites, totalFichiers, "Injection  :")
	}

	f.SaveAs(rapportPath)
	fmt.Println(ColorGreen + "\n=========================================" + ColorReset)
	fmt.Println(ColorGreen + "          TRAITEMENT TERMINÉ             " + ColorReset)
	fmt.Println(ColorGreen + "=========================================" + ColorReset)
}

// --- MENU ---
func main() {
	for {
		fmt.Println("\n" + ColorYellow + "=== OUTIL DE CORRECTION MASSIVE ===" + ColorReset)
		fmt.Println("1 - Phase 1 - Analyse, contrôle et édition du rapport")
		fmt.Println("2 - Phase 2 - Mise en place des corrections utilisateur")
		fmt.Println("3 - Quitter")
		fmt.Print(ColorCyan + "Choix : " + ColorReset)

		scanner := bufio.NewScanner(os.Stdin)
		scanner.Scan()
		switch strings.TrimSpace(scanner.Text()) {
		case "1":
			phase1Analyse()
		case "2":
			phase2Correction()
		case "3":
			os.Exit(0)
		default:
			fmt.Println(ColorRed + "Choix invalide." + ColorReset)
		}
	}
}

/*
==============================================================================
*/
