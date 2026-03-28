/*
==============================================================================
Titre : Moteur d'Audit et Correction Linguistique Massif (Multi-Formats)
Auteur : Digixtp
Version : 8.1.0 (Prompt Défensif & Monitoring Absolu)
Note introductive : Ce script orchestre la correction de textes via un LLM local.
[MÀJ V8.1.0] : Refonte du prompt pour contrer les hallucinations typographiques
des petits modèles (interdiction de modifier la casse initiale ou les apostrophes).
Correction de l'algorithme de traçabilité : le fichier de log se génère désormais
correctement si aucune erreur finale n'est retenue, permettant un audit réel
du comportement du modèle face aux vraies coquilles (ex: sas -> sans).
==============================================================================
*/
package main

import (
	"bufio"
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"syscall"
	"time"
	"unicode"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

// --- CONSTANTES ANSI ---
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
	Temperature   float64 `json:"temperature"`
	RepeatPenalty float64 `json:"repeat_penalty"`
	NumCtx        int     `json:"num_ctx"`
	NumPredict    int     `json:"num_predict"`
}

type LLMResponse struct {
	Response string `json:"response"`
}

type LigneSource struct {
	Coordonnees string
	Texte       string
}

type BlockMajeur struct {
	Index  int
	Lignes []LigneSource
	Chemin string
}

type StateDB struct {
	Processed map[string]bool `json:"processed"`
}

type ErreurDetecteePropre struct {
	PhraseOriginale string
	MotFaux         string
	MotCorrige      string
}

type ActionUnitaire struct {
	MotFaux    string
	MotCorrige string
	LigneExcel int
}

type RemplacementAgrege struct {
	Original        string
	CorrectionFinal string
	LignesExcel     []int
}

// --- CLIENT HTTP OPTIMISÉ ---
var httpClient = &http.Client{
	Transport: &http.Transport{
		MaxIdleConns:        100,
		MaxIdleConnsPerHost: 100,
		IdleConnTimeout:     90 * time.Second,
	},
}

var stopRequested int32

// --- INITIALISATION SYSTÈME ---
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

// --- UTILITAIRES SYSTÈME ET SÉCURITÉ ---
func getDesktopPath() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return "."
	}
	return filepath.Join(home, "Desktop")
}

func bypassMaxPath(p string) string {
	abs, err := filepath.Abs(p)
	if err != nil {
		return p
	}
	if !strings.HasPrefix(abs, `\\?\`) {
		return `\\?\` + abs
	}
	return abs
}

func safeDelete(filePath string) {
	trashDir := filepath.Join(getDesktopPath(), "fichier à supprimer")
	os.MkdirAll(trashDir, 0755)

	base := filepath.Base(filePath)
	dest := filepath.Join(trashDir, fmt.Sprintf("%d_corrupted_%s", time.Now().Unix(), base))

	err := os.Rename(filePath, dest)
	if err != nil && !os.IsNotExist(err) {
		fmt.Printf("\n%s[Avertissement] Impossible de sécuriser le fichier : %v%s\n", ColorYellow, err, ColorReset)
	}
}

func purgerMemoireAutomatique() {
	statePath := filepath.Join(getDesktopPath(), "Audit_State.json")
	if _, err := os.Stat(bypassMaxPath(statePath)); err == nil {
		safeDelete(bypassMaxPath(statePath))
	}
}

func loadState(path string) StateDB {
	db := StateDB{Processed: make(map[string]bool)}
	data, err := os.ReadFile(bypassMaxPath(path))
	if err == nil {
		errUnmarshall := json.Unmarshal(data, &db)
		if errUnmarshall != nil {
			fmt.Println(ColorRed + "[Alerte] Fichier d'état corrompu. Sécurisation..." + ColorReset)
			safeDelete(bypassMaxPath(path))
			return StateDB{Processed: make(map[string]bool)}
		}
	}
	return db
}

func saveState(path string, db StateDB) {
	data, err := json.MarshalIndent(db, "", "  ")
	if err == nil {
		os.WriteFile(bypassMaxPath(path), data, 0644)
	}
}

func logDebugLLM(raw string, parseErr string) {
	logPath := filepath.Join(getDesktopPath(), "LLM_Debug_Crash.log")
	f, err := os.OpenFile(logPath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err == nil {
		defer f.Close()
		timestamp := time.Now().Format("2006-01-02 15:04:05")
		f.WriteString(fmt.Sprintf("\n[%s] === ERREUR PARSING JSON ===\n", timestamp))
		f.WriteString("Erreur Go : " + parseErr + "\n")
		f.WriteString("Réponse LLM Brute :\n" + raw + "\n===========================\n")
	}
}

func logTraceZeroFaute(texteSoumis string, reponseLLM string) {
	logPath := filepath.Join(getDesktopPath(), "LLM_Trace_Zero_Faute.log")
	f, err := os.OpenFile(logPath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err == nil {
		defer f.Close()
		timestamp := time.Now().Format("2006-01-02 15:04:05")
		f.WriteString(fmt.Sprintf("\n[%s] === TRACE D'ANALYSE (0 Faute Retenue) ===\n", timestamp))
		f.WriteString("--- TEXTE SOUMIS AU LLM ---\n" + texteSoumis + "\n")
		f.WriteString("--- RÉPONSE BRUTE DU LLM ---\n" + reponseLLM + "\n===========================\n")
	}
}

func verifierVerrouFichier(chemin string) error {
	file, err := os.OpenFile(bypassMaxPath(chemin), os.O_RDWR, 0666)
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

func afficherBarreProgression(iteration, total int, prefix string) {
	if atomic.LoadInt32(&stopRequested) == 1 {
		return
	}
	length := 40
	percent := float64(iteration) / float64(total) * 100
	filled := int(float64(length) * float64(iteration) / float64(total))
	bar := strings.Repeat("█", filled) + strings.Repeat("-", length-filled)
	fmt.Printf("\r%s%s%s |%s| %.1f%% (%d/%d)", ColorBlue, prefix, ColorReset, bar, percent, iteration, total)
	if iteration == total {
		fmt.Println()
	}
}

func parseFlexString(val interface{}) string {
	if val == nil {
		return ""
	}
	switch v := val.(type) {
	case string:
		return v
	case []interface{}:
		var strs []string
		for _, s := range v {
			if strVal, ok := s.(string); ok {
				strs = append(strs, strVal)
			} else {
				strs = append(strs, fmt.Sprintf("%v", s))
			}
		}
		return strings.Join(strs, ", ")
	default:
		return fmt.Sprintf("%v", v)
	}
}

func estTexteValidePourCorrection(texte string) bool {
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
	return true
}

// --- PARSING FICHIERS OCR ---
func parserFichierOCR(ocrPath string) ([]BlockMajeur, error) {
	bytesData, err := os.ReadFile(bypassMaxPath(ocrPath))
	if err != nil {
		return nil, err
	}
	contenu := string(bytesData)

	cheminNatif := ""
	lignes := strings.Split(contenu, "\n")
	for _, ligne := range lignes {
		if strings.HasPrefix(ligne, "Chemin_complet__") {
			cheminNatif = strings.TrimSpace(strings.TrimPrefix(ligne, "Chemin_complet__"))
			cheminNatif = strings.TrimPrefix(cheminNatif, `\\?\`)
			break
		}
	}

	if cheminNatif == "" {
		return nil, fmt.Errorf("chemin natif introuvable")
	}

	debIdx := strings.Index(contenu, "__début_du_corpus___")
	finIdx := strings.Index(contenu, "__fin_du_corpus___")
	if debIdx == -1 || finIdx == -1 || debIdx >= finIdx {
		return nil, fmt.Errorf("balises corpus introuvables")
	}

	corpus := contenu[debIdx+len("__début_du_corpus___") : finIdx]
	lignesCorpus := strings.Split(corpus, "\n")
	reTags := regexp.MustCompile(`^((?:\[.*?\])+)\s*(.*)`)

	var blocs []BlockMajeur
	indexBloc := 0
	blocCourant := BlockMajeur{Chemin: cheminNatif, Index: indexBloc}
	var charCount int

	for _, ligne := range lignesCorpus {
		ligne = strings.TrimSpace(ligne)
		if ligne == "" || ligne == "[Aucun contenu textuel extrait de l'image]" {
			continue
		}

		matches := reTags.FindStringSubmatch(ligne)
		currentTags := "[Sans_Coordonnee]"
		textPur := ligne

		if len(matches) == 3 {
			currentTags = strings.TrimSpace(matches[1])
			textPur = strings.TrimSpace(matches[2])
		}

		if estTexteValidePourCorrection(textPur) {
			if charCount+len(textPur) > 1000 && len(blocCourant.Lignes) > 0 {
				blocs = append(blocs, blocCourant)
				indexBloc++
				blocCourant = BlockMajeur{Chemin: cheminNatif, Index: indexBloc}
				charCount = 0
			}
			blocCourant.Lignes = append(blocCourant.Lignes, LigneSource{Coordonnees: currentTags, Texte: textPur})
			charCount += len(textPur)
		}
	}

	if len(blocCourant.Lignes) > 0 {
		blocs = append(blocs, blocCourant)
	}

	return blocs, nil
}

// --- MOTEUR LLM ---
func interrogerLLM(texte string) []ErreurDetecteePropre {
	url := "http://localhost:11434/api/generate"

	// Ingénierie de prompt défensive ciblant spécifiquement les erreurs de Qwen 1.5b
	prompt := fmt.Sprintf(`Tu es un correcteur orthographique d'élite.
Ta mission est d'identifier UNIQUEMENT les vraies coquilles (erreurs de frappe évidentes, comme "sas" au lieu de "sans") ou les erreurs grammaticales claires.

RÈGLES DE PÉNALITÉ ABSOLUES :
1. INTERDICTION FORMELLE de modifier les majuscules en début de phrase ou de mot (ex: "L'oubli" DOIT rester "L'oubli").
2. INTERDICTION FORMELLE de modifier la ponctuation ou de casser les apostrophes (ex: "L'oubli" ne devient jamais "leoubli").
3. Si le texte ne contient aucune véritable faute de frappe, renvoie EXACTEMENT et UNIQUEMENT : []
4. Réponds UNIQUEMENT par un tableau JSON valide.

Format strict attendu (crée un objet par faute trouvée) :
[
  {
    "phrase_originale": "Texte exact fourni.",
    "mot_faux": "sas",
    "mot_corrige": "sans"
  }
]

Texte à analyser :
%s`, texte)

	reqBody := LLMRequest{
		Model:  "qwen2.5:1.5b",
		Prompt: prompt,
		Stream: false,
		Format: "json",
		Options: LLMOption{
			Temperature:   0.0,
			RepeatPenalty: 1.15,
			NumCtx:        2048,
			NumPredict:    1024,
		},
	}

	jsonData, _ := json.Marshal(reqBody)
	ctx, cancel := context.WithTimeout(context.Background(), 180*time.Second)
	defer cancel()

	req, err := http.NewRequestWithContext(ctx, "POST", url, bytes.NewBuffer(jsonData))
	if err != nil {
		return nil
	}
	req.Header.Set("Content-Type", "application/json")

	resp, err := httpClient.Do(req)
	if err != nil {
		return nil
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var llmResp LLMResponse
	json.Unmarshal(body, &llmResp)

	rawOutput := strings.TrimSpace(llmResp.Response)

	if strings.HasPrefix(rawOutput, "```json") {
		rawOutput = strings.TrimPrefix(rawOutput, "```json")
	} else if strings.HasPrefix(rawOutput, "```") {
		rawOutput = strings.TrimPrefix(rawOutput, "```")
	}
	rawOutput = strings.TrimSuffix(rawOutput, "```")
	cleanJSON := strings.TrimSpace(rawOutput)

	if strings.HasPrefix(cleanJSON, "{") && strings.HasSuffix(cleanJSON, "}") {
		cleanJSON = "[" + cleanJSON + "]"
	}

	var erreursDynamiques []map[string]interface{}
	errUnmarshall := json.Unmarshal([]byte(cleanJSON), &erreursDynamiques)

	if errUnmarshall != nil {
		if cleanJSON != "[]" && cleanJSON != "" {
			logDebugLLM(cleanJSON, errUnmarshall.Error())
		}
		return nil
	}

	var erreursPropres []ErreurDetecteePropre

	for _, errMap := range erreursDynamiques {
		phraseOrig := parseFlexString(errMap["phrase_originale"])
		mFaux := parseFlexString(errMap["mot_faux"])
		mCorr := parseFlexString(errMap["mot_corrige"])

		if phraseOrig != "" && mFaux != "" && mCorr != "" {
			// Validation de sécurité : Empêcher le modèle de s'auto-corriger sur la casse du premier mot
			if strings.EqualFold(mFaux, mCorr) {
				continue
			}

			erreursPropres = append(erreursPropres, ErreurDetecteePropre{
				PhraseOriginale: phraseOrig,
				MotFaux:         mFaux,
				MotCorrige:      mCorr,
			})
		}
	}

	// TRACE ABSOLUE : Si aucune erreur VALIDE n'a été retenue, on log la réponse brute
	if len(erreursPropres) == 0 && len(texte) > 3 {
		logTraceZeroFaute(texte, cleanJSON)
	}

	return erreursPropres
}

// --- COLORISATION RICH TEXT ---
func getRichTextRuns(phrase string, targetWord string, colorHex string) []excelize.RichTextRun {
	if targetWord == "" || strings.TrimSpace(targetWord) == "aucun" || strings.TrimSpace(phrase) == "" {
		return []excelize.RichTextRun{{Text: phrase}}
	}

	runesPhrase := []rune(phrase)
	phraseLower := []rune(strings.ToLower(phrase))
	colors := make([]bool, len(runesPhrase))

	tLower := []rune(strings.ToLower(strings.TrimSpace(targetWord)))

	if len(tLower) > 0 {
		for i := 0; i <= len(phraseLower)-len(tLower); i++ {
			match := true
			for j := range tLower {
				if phraseLower[i+j] != tLower[j] {
					match = false
					break
				}
			}
			if match {
				for j := range tLower {
					colors[i+j] = true
				}
			}
		}
	}

	var runs []excelize.RichTextRun
	if len(runesPhrase) == 0 {
		return runs
	}

	runStart := 0
	currentIsColor := colors[0]

	for i := 1; i < len(runesPhrase); i++ {
		if colors[i] != currentIsColor {
			if currentIsColor {
				runs = append(runs, excelize.RichTextRun{Text: string(runesPhrase[runStart:i]), Font: &excelize.Font{Color: colorHex, Bold: true}})
			} else {
				runs = append(runs, excelize.RichTextRun{Text: string(runesPhrase[runStart:i])})
			}
			runStart = i
			currentIsColor = colors[i]
		}
	}

	if currentIsColor {
		runs = append(runs, excelize.RichTextRun{Text: string(runesPhrase[runStart:]), Font: &excelize.Font{Color: colorHex, Bold: true}})
	} else {
		runs = append(runs, excelize.RichTextRun{Text: string(runesPhrase[runStart:])})
	}

	return runs
}

// --- MOTEUR D'INJECTION COM ---
func corrigerOfficeCOM_Batch(chemin string, remplacements []RemplacementAgrege, typeFichier string) error {
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
			corrClean := strings.ReplaceAll(remp.CorrectionFinal, "'", "''")
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
			corrClean := strings.ReplaceAll(remp.CorrectionFinal, "'", "''")
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

func corrigerExcel_Batch(chemin string, remplacements []RemplacementAgrege) error {
	xl, err := excelize.OpenFile(bypassMaxPath(chemin))
	if err != nil {
		return err
	}
	defer xl.Close()

	for _, sheetName := range xl.GetSheetList() {
		rows, _ := xl.GetRows(sheetName)
		for rowIndex, row := range rows {
			for colIndex, cellValue := range row {
				for _, remp := range remplacements {
					if strings.Contains(cellValue, remp.Original) {
						newVal := strings.Replace(cellValue, remp.Original, remp.CorrectionFinal, 1)
						cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
						xl.SetCellValue(sheetName, cellName, newVal)
						cellValue = newVal
					}
				}
			}
		}
	}
	return xl.Save()
}

// --- FLUX PRINCIPAL (PHASE 1) ---
func phase1Analyse() {
	purgerMemoireAutomatique()

	c := make(chan os.Signal, 1)
	signal.Notify(c, os.Interrupt, syscall.SIGTERM)
	go func() {
		<-c
		fmt.Println("\n\n" + ColorYellow + "[Arrêt Demandé] Sauvegarde de la progression en cours..." + ColorReset)
		atomic.StoreInt32(&stopRequested, 1)
	}()

	fmt.Print(ColorCyan + "Chemin absolu du dossier database_OCR à analyser : " + ColorReset)
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
	dossierCible := strings.TrimSpace(scanner.Text())

	if _, err := os.Stat(bypassMaxPath(dossierCible)); os.IsNotExist(err) {
		fmt.Println(ColorRed + "Erreur : Le dossier n'existe pas." + ColorReset)
		return
	}

	desktopPath := getDesktopPath()
	rapportPath := filepath.Join(desktopPath, "Audit_Corrections.xlsx")
	statePath := filepath.Join(desktopPath, "Audit_State.json")

	stateDB := loadState(statePath)
	var f *excelize.File
	currentRow := 6
	sheet := "Audit_Corrections"

	if _, err := os.Stat(bypassMaxPath(rapportPath)); err == nil {
		fmt.Println(ColorGreen + "[Info] Fichier d'audit existant détecté. Mode Reprise activé." + ColorReset)
		f, _ = excelize.OpenFile(rapportPath)
		rows, _ := f.GetRows(sheet)
		if len(rows) >= 5 {
			currentRow = len(rows) + 1
		}
	} else {
		f = excelize.NewFile()
		f.SetSheetName("Sheet1", sheet)

		f.SetCellValue(sheet, "A1", "RAPPORT D'AUDIT LLM (Piloté par Data OCR)")
		f.SetCellValue(sheet, "A2", "Note explicative : Ce fichier centralise les fautes isolées par le modèle.")
		f.SetCellValue(sheet, "A3", "Les données commencent à la ligne 5 pour respecter le standard de présentation.")

		f.MergeCell(sheet, "A1", "H1")
		f.MergeCell(sheet, "A2", "H2")
		f.MergeCell(sheet, "A3", "H3")

		styleIntro, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Color: "#333333"},
			Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		})
		f.SetCellStyle(sheet, "A1", "H3", styleIntro)

		styleData, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{WrapText: true, Vertical: "top"},
		})
		f.SetColStyle(sheet, "A:H", styleData)

		styleTitre, _ := f.NewStyle(&excelize.Style{
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"#000080"}, Pattern: 1},
			Font:      &excelize.Font{Bold: true, Color: "#FFFFFF"},
			Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		})

		enTetes := []string{"Coordonnées", "Chemin du fichier (Natif)", "Texte de la balise", "Aperçu (Correction Unitaire)", "Statut d'injection", "Validation (v/r)", "Tech_Faux", "Tech_Corrige"}
		for i, colName := range enTetes {
			cell := string(rune('A'+i)) + "5"
			f.SetCellValue(sheet, cell, colName)
			f.SetCellStyle(sheet, cell, cell, styleTitre)
		}

		f.SetColWidth(sheet, "A", "A", 25)
		f.SetColWidth(sheet, "B", "B", 35)
		f.SetColWidth(sheet, "C", "C", 55)
		f.SetColWidth(sheet, "D", "D", 55)
		f.SetColWidth(sheet, "E", "E", 20)
		f.SetColWidth(sheet, "F", "F", 20)

		f.SetColVisible(sheet, "G", false)
		f.SetColVisible(sheet, "H", false)
	}

	var fichiersOCR []string
	filepath.Walk(bypassMaxPath(dossierCible), func(path string, info os.FileInfo, err error) error {
		if err == nil && !info.IsDir() && strings.HasSuffix(path, ".ocr.txt") {
			fichiersOCR = append(fichiersOCR, path)
		}
		return nil
	})

	if len(fichiersOCR) == 0 {
		fmt.Println(ColorYellow + "Aucun fichier .ocr.txt trouvé." + ColorReset)
		return
	}

	var tousLesBlocs []BlockMajeur
	for _, path := range fichiersOCR {
		blocs, _ := parserFichierOCR(path)
		for _, b := range blocs {
			tousLesBlocs = append(tousLesBlocs, b)
		}
	}
	totalBlocsGlobal := len(tousLesBlocs)

	fmt.Printf("%s[Step] Traitement par IA (%d blocs) (Ctrl+C pour arrêter proprement)%s\n", ColorBlue, totalBlocsGlobal, ColorReset)

	var mu sync.Mutex
	var wg sync.WaitGroup
	semaphore := make(chan struct{}, 5)

	blocsTraitesGlobal := 0
	erreursTrouveesGlobal := 0

	for _, bloc := range tousLesBlocs {
		if atomic.LoadInt32(&stopRequested) == 1 {
			break
		}

		blockKey := fmt.Sprintf("%s_bloc_%d", filepath.Base(bloc.Chemin), bloc.Index)
		if stateDB.Processed[blockKey] {
			blocsTraitesGlobal++
			continue
		}

		wg.Add(1)
		semaphore <- struct{}{}

		go func(b BlockMajeur, bKey string) {
			defer wg.Done()
			defer func() { <-semaphore }()

			for _, ls := range b.Lignes {
				textePropre := strings.TrimSpace(ls.Texte)
				if len(textePropre) < 3 {
					continue
				}

				erreurs := interrogerLLM(textePropre)

				mu.Lock()
				if len(erreurs) > 0 {
					for _, errDetect := range erreurs {
						if errDetect.MotFaux != "" && strings.Contains(textePropre, errDetect.MotFaux) {

							phraseUnitaire := strings.Replace(textePropre, errDetect.MotFaux, errDetect.MotCorrige, 1)

							f.SetCellValue(sheet, fmt.Sprintf("A%d", currentRow), ls.Coordonnees)
							f.SetCellValue(sheet, fmt.Sprintf("B%d", currentRow), b.Chemin)

							runsOrig := getRichTextRuns(textePropre, errDetect.MotFaux, "FF0000")
							f.SetCellRichText(sheet, fmt.Sprintf("C%d", currentRow), runsOrig)

							runsCorr := getRichTextRuns(phraseUnitaire, errDetect.MotCorrige, "00B050")
							f.SetCellRichText(sheet, fmt.Sprintf("D%d", currentRow), runsCorr)

							f.SetCellValue(sheet, fmt.Sprintf("E%d", currentRow), "En attente Validation")

							f.SetCellValue(sheet, fmt.Sprintf("G%d", currentRow), errDetect.MotFaux)
							f.SetCellValue(sheet, fmt.Sprintf("H%d", currentRow), errDetect.MotCorrige)

							currentRow++
							erreursTrouveesGlobal++
						}
					}
				}
				mu.Unlock()
			}

			mu.Lock()
			stateDB.Processed[bKey] = true
			blocsTraitesGlobal++
			afficherBarreProgression(blocsTraitesGlobal, totalBlocsGlobal, "Analyse LLM")

			if blocsTraitesGlobal%5 == 0 {
				saveState(statePath, stateDB)
			}
			mu.Unlock()

		}(bloc, blockKey)
	}

	wg.Wait()
	fmt.Println()

	saveState(statePath, stateDB)
	f.SaveAs(rapportPath)

	fmt.Println(ColorGreen + "\n=========================================" + ColorReset)
	if atomic.LoadInt32(&stopRequested) == 1 {
		fmt.Println(ColorYellow + "        ARRÊT PRÉMATURÉ SAUVEGARDÉ       " + ColorReset)
	} else {
		fmt.Println(ColorGreen + "          TRAITEMENT TERMINÉ             " + ColorReset)
	}
	fmt.Printf("Fautes isolées dans le rapport : %d\n", erreursTrouveesGlobal)
	fmt.Println(ColorGreen + "=========================================" + ColorReset)
}

// --- FLUX PRINCIPAL (PHASE 2) ---
func phase2Correction() {
	purgerMemoireAutomatique()

	c := make(chan os.Signal, 1)
	signal.Notify(c, os.Interrupt, syscall.SIGTERM)
	go func() {
		<-c
		fmt.Println("\n\n" + ColorYellow + "[Arrêt Demandé] Terminaison et sauvegarde..." + ColorReset)
		atomic.StoreInt32(&stopRequested, 1)
	}()

	rapportPath := filepath.Join(getDesktopPath(), "Audit_Corrections.xlsx")
	f, err := excelize.OpenFile(bypassMaxPath(rapportPath))
	if err != nil {
		fmt.Println(ColorRed + "Erreur : Rapport d'audit introuvable sur le bureau." + ColorReset)
		return
	}
	defer f.Close()

	rows, _ := f.GetRows("Audit_Corrections")

	planDActionBrut := make(map[string]map[string][]ActionUnitaire)

	for i := 5; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 8 {
			continue
		}

		cheminNatif := row[1]
		original := row[2]
		statutActuel := row[4]
		validation := strings.ToLower(strings.TrimSpace(row[5]))
		motFaux := row[6]
		motCorrige := row[7]

		if statutActuel == "Validé et Injecté" || statutActuel == "Ignoré (Refusé)" {
			continue
		}

		if validation != "v" {
			f.SetCellValue("Audit_Corrections", fmt.Sprintf("E%d", i+1), "Ignoré (Refusé)")
			continue
		}

		if planDActionBrut[cheminNatif] == nil {
			planDActionBrut[cheminNatif] = make(map[string][]ActionUnitaire)
		}

		planDActionBrut[cheminNatif][original] = append(planDActionBrut[cheminNatif][original], ActionUnitaire{
			MotFaux:    motFaux,
			MotCorrige: motCorrige,
			LigneExcel: i + 1,
		})
	}

	totalFichiers := len(planDActionBrut)
	if totalFichiers == 0 {
		fmt.Println(ColorYellow + "Aucune nouvelle correction validée ('v') à appliquer." + ColorReset)
		f.SaveAs(rapportPath)
		return
	}

	fmt.Printf("%s[Step] Injection dans %d fichiers (Ctrl+C pour interrompre)...%s\n", ColorBlue, totalFichiers, ColorReset)
	fichiersTraites := 0

	for chemin, mapPhrases := range planDActionBrut {
		if atomic.LoadInt32(&stopRequested) == 1 {
			break
		}

		var listeAgregee []RemplacementAgrege
		for phraseOrig, actions := range mapPhrases {
			phraseEvolutive := phraseOrig
			var lignesConcernees []int

			for _, act := range actions {
				phraseEvolutive = strings.Replace(phraseEvolutive, act.MotFaux, act.MotCorrige, 1)
				lignesConcernees = append(lignesConcernees, act.LigneExcel)
			}

			listeAgregee = append(listeAgregee, RemplacementAgrege{
				Original:        phraseOrig,
				CorrectionFinal: phraseEvolutive,
				LignesExcel:     lignesConcernees,
			})
		}

		var errInjection error
		if errVerrou := verifierVerrouFichier(chemin); errVerrou != nil {
			errInjection = errVerrou
		} else {
			ext := strings.ToLower(filepath.Ext(chemin))
			switch ext {
			case ".txt", ".md":
				contentBytes, err := os.ReadFile(bypassMaxPath(chemin))
				if err != nil {
					errInjection = err
				} else {
					newContent := string(contentBytes)
					for _, remp := range listeAgregee {
						newContent = strings.Replace(newContent, remp.Original, remp.CorrectionFinal, 1)
					}
					attendreDisqueOptimal(60)
					errInjection = os.WriteFile(bypassMaxPath(chemin), []byte(newContent), 0644)
				}
			case ".docx", ".pptx":
				attendreDisqueOptimal(60)
				errInjection = corrigerOfficeCOM_Batch(chemin, listeAgregee, ext)
			case ".xlsx", ".xlsm":
				attendreDisqueOptimal(60)
				errInjection = corrigerExcel_Batch(chemin, listeAgregee)
			}
		}

		for _, remp := range listeAgregee {
			for _, ligneExcel := range remp.LignesExcel {
				celluleStatut := fmt.Sprintf("E%d", ligneExcel)
				if errInjection != nil {
					styleErr, _ := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#f8d7da"}, Pattern: 1}, Font: &excelize.Font{Color: "#721c24"}})
					f.SetCellValue("Audit_Corrections", celluleStatut, "Erreur : "+errInjection.Error())
					f.SetCellStyle("Audit_Corrections", celluleStatut, celluleStatut, styleErr)
				} else {
					styleOk, _ := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#d4edda"}, Pattern: 1}, Font: &excelize.Font{Color: "#155724"}})
					f.SetCellValue("Audit_Corrections", celluleStatut, "Validé et Injecté")
					f.SetCellStyle("Audit_Corrections", celluleStatut, celluleStatut, styleOk)
				}
			}
		}

		fichiersTraites++
		afficherBarreProgression(fichiersTraites, totalFichiers, "Injection Physique")
		f.SaveAs(rapportPath)
	}

	fmt.Println(ColorGreen + "\n\n=========================================" + ColorReset)
	if atomic.LoadInt32(&stopRequested) == 1 {
		fmt.Println(ColorYellow + "      INJECTION INTERROMPUE PROPREMENT   " + ColorReset)
	} else {
		fmt.Println(ColorGreen + "          INJECTION TERMINÉE             " + ColorReset)
	}
	fmt.Println(ColorGreen + "=========================================" + ColorReset)
}

// --- MENU ---
func main() {
	for {
		atomic.StoreInt32(&stopRequested, 0)
		fmt.Println("\n" + ColorYellow + "=== OUTIL DE CORRECTION MASSIVE ===" + ColorReset)
		fmt.Println("1 - Phase 1 - Analyser la base de données OCR")
		fmt.Println("2 - Phase 2 - Injecter les corrections dans les fichiers natifs")
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
