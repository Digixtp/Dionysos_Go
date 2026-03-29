/*
==============================================================================
Titre : Moteur d'Audit Linguistique Massif (Mode Détection)
Auteur : Digixtp
Version : 14.0.0 (Golden Master - Archive)
Note introductive : Script d'orchestration pour l'audit de textes via Ollama.
Cette version est sanctuarisée. L'architecture Go (Batching, Diffing, Regex,
Export Excel RichText) est aboutie. En attente d'un LLM disposant de >7B
paramètres pour garantir un respect strict du format JSON sans hallucination.
==============================================================================
*/
package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"os/signal"
	"path/filepath"
	"regexp"
	"strings"
	"sync"
	"sync/atomic"
	"syscall"
	"time"
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

// --- CLIENT HTTP ---
var httpClient = &http.Client{
	Timeout: 300 * time.Second,
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
	home, _ := os.UserHomeDir()
	return filepath.Join(home, "Desktop")
}

func bypassMaxPath(p string) string {
	abs, _ := filepath.Abs(p)
	if !strings.HasPrefix(abs, `\\?\`) {
		return `\\?\` + abs
	}
	return abs
}

func safeDelete(src string) {
	trashDir := filepath.Join(getDesktopPath(), "fichier à supprimer")
	os.MkdirAll(trashDir, 0755)
	ts := time.Now().Format("20060102_150405")
	dest := filepath.Join(trashDir, ts+"_"+filepath.Base(src))
	os.Rename(src, dest)
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
		json.Unmarshal(data, &db)
	}
	return db
}

func saveState(path string, db StateDB) {
	data, _ := json.MarshalIndent(db, "", "  ")
	os.WriteFile(bypassMaxPath(path), data, 0644)
}

func logTraceZeroFaute(lot string, resp string) {
	logPath := filepath.Join(getDesktopPath(), "LLM_Trace_Zero_Faute.log")
	f, err := os.OpenFile(logPath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err == nil {
		defer f.Close()
		f.WriteString(fmt.Sprintf("\n[%s] --- ANALYSE LOT ---\nLot: %s\nReponse:\n%s\n", time.Now().Format("15:04:05"), lot, resp))
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
	default:
		return fmt.Sprintf("%v", v)
	}
}

// --- MOTEUR DE DIFFING ---
func findExactDiff(orig, corr string) (string, string) {
	origR := []rune(orig)
	corrR := []rune(corr)

	start := 0
	minLen := len(origR)
	if len(corrR) < minLen {
		minLen = len(corrR)
	}

	for start < minLen && origR[start] == corrR[start] {
		start++
	}

	for start > 0 && origR[start-1] != ' ' && origR[start-1] != '\'' && origR[start-1] != '-' {
		start--
	}

	endOrig := len(origR)
	endCorr := len(corrR)

	for endOrig > start && endCorr > start && origR[endOrig-1] == corrR[endCorr-1] {
		endOrig--
		endCorr--
	}

	for endOrig < len(origR) && origR[endOrig] != ' ' && origR[endOrig] != '.' && origR[endOrig] != ',' && origR[endOrig] != ';' && origR[endOrig] != ':' {
		endOrig++
	}
	for endCorr < len(corrR) && corrR[endCorr] != ' ' && corrR[endCorr] != '.' && corrR[endCorr] != ',' && corrR[endCorr] != ';' && corrR[endCorr] != ':' {
		endCorr++
	}

	if start >= endOrig || start >= endCorr {
		return orig, corr
	}

	return strings.TrimSpace(string(origR[start:endOrig])), strings.TrimSpace(string(corrR[start:endCorr]))
}

// --- COLORISATION RICH TEXT ---
func getRichTextRuns(phrase string, targetWord string, colorHex string) []excelize.RichTextRun {
	if targetWord == "" || phrase == "" {
		return []excelize.RichTextRun{{Text: phrase}}
	}

	re, err := regexp.Compile(`(?i)` + regexp.QuoteMeta(targetWord))
	if err != nil {
		return []excelize.RichTextRun{{Text: phrase}}
	}

	locs := re.FindAllStringIndex(phrase, -1)
	if len(locs) == 0 {
		return []excelize.RichTextRun{{Text: phrase}}
	}

	var runs []excelize.RichTextRun
	lastIdx := 0

	for _, loc := range locs {
		if loc[0] > lastIdx {
			runs = append(runs, excelize.RichTextRun{Text: phrase[lastIdx:loc[0]]})
		}
		actualWord := phrase[loc[0]:loc[1]]
		runs = append(runs, excelize.RichTextRun{
			Text: actualWord,
			Font: &excelize.Font{Color: colorHex, Bold: true},
		})
		lastIdx = loc[1]
	}
	if lastIdx < len(phrase) {
		runs = append(runs, excelize.RichTextRun{Text: phrase[lastIdx:]})
	}
	return runs
}

// --- MOTEUR OLLAMA ---
func validerPreRequisLLM() error {
	client := http.Client{Timeout: 3 * time.Second}
	resp, err := client.Get("http://localhost:11434/api/tags")
	if err != nil {
		return fmt.Errorf("Ollama injoignable. Le service doit être démarré")
	}
	defer resp.Body.Close()
	return nil
}

func interrogerLLM_Detection(phrases []string) []ErreurDetecteePropre {
	if len(phrases) == 0 {
		return nil
	}

	lotJSON, _ := json.Marshal(phrases)
	url := "http://localhost:11434/api/generate"

	prompt := fmt.Sprintf(`Analyse ces phrases et trouve UNIQUEMENT les fautes d'orthographe ou de grammaire.
RÈGLES D'OR ABSOLUES :
1. INTERDICTION D'UTILISER DES SYNONYMES. 
2. Si une phrase est correcte, ignore-la.
3. Tu DOIS renvoyer un objet JSON contenant un unique tableau nommé "corrections".

FORMAT EXACT EXIGÉ :
{
  "corrections": [
    {
      "phrase": "la phrase entière avec sa faute",
      "faux": "le mot erroné exact",
      "corrige": "la correction orthographique"
    }
  ]
}

PHRASES À TRAITER : 
%s`, string(lotJSON))

	reqBody := LLMRequest{
		Model: "qwen2.5:1.5b", Prompt: prompt, Stream: false, Format: "json",
		Options: LLMOption{Temperature: 0.0, RepeatPenalty: 1.0, NumCtx: 4096, NumPredict: 2048},
	}

	jsonData, _ := json.Marshal(reqBody)
	resp, err := httpClient.Post(url, "application/json", bytes.NewBuffer(jsonData))
	if err != nil {
		return nil
	}
	defer resp.Body.Close()

	var llmResp LLMResponse
	json.NewDecoder(resp.Body).Decode(&llmResp)

	raw := strings.TrimSpace(llmResp.Response)
	raw = strings.TrimPrefix(raw, "```json")
	raw = strings.TrimSuffix(raw, "```")
	raw = strings.TrimSpace(raw)

	var parsedData map[string]interface{}
	if err := json.Unmarshal([]byte(raw), &parsedData); err != nil {
		return nil
	}

	rawCorrections, ok := parsedData["corrections"].([]interface{})
	if !ok {
		return nil
	}

	var results []ErreurDetecteePropre

	for _, item := range rawCorrections {
		obj, isMap := item.(map[string]interface{})
		if !isMap {
			continue
		}

		p := parseFlexString(obj["phrase"])
		f := parseFlexString(obj["faux"])
		c := parseFlexString(obj["corrige"])

		if f == "" || c == "" || p == "" {
			continue
		}
		if strings.EqualFold(f, c) {
			continue
		}
		if len(c) > len(f)*3 || len(f) > len(c)*3 {
			continue
		}

		phraseTrouvee := ""
		for _, phraseSource := range phrases {
			if strings.Contains(strings.ToLower(phraseSource), strings.ToLower(f)) || strings.Contains(phraseSource, p) {
				phraseTrouvee = phraseSource
				break
			}
		}

		if phraseTrouvee != "" {
			results = append(results, ErreurDetecteePropre{PhraseOriginale: phraseTrouvee, MotFaux: f, MotCorrige: c})
		}
	}

	if len(results) == 0 {
		logTraceZeroFaute(string(lotJSON), raw)
	}
	return results
}

// --- PARSING FICHIERS OCR ---
func parserFichierOCR(ocrPath string) ([]BlockMajeur, error) {
	bytesData, err := os.ReadFile(bypassMaxPath(ocrPath))
	if err != nil {
		return nil, fmt.Errorf("lecture impossible")
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
		return nil, fmt.Errorf("balise 'Chemin_complet__' introuvable")
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

		if len(textPur) > 4 && strings.ContainsAny(textPur, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ") {
			if len(blocCourant.Lignes) >= 6 {
				blocs = append(blocs, blocCourant)
				indexBloc++
				blocCourant = BlockMajeur{Chemin: cheminNatif, Index: indexBloc}
			}
			blocCourant.Lignes = append(blocCourant.Lignes, LigneSource{Coordonnees: currentTags, Texte: textPur})
		}
	}
	if len(blocCourant.Lignes) > 0 {
		blocs = append(blocs, blocCourant)
	}

	return blocs, nil
}

// --- LOGIQUE MÉTIER ---
func lancerAuditDetection() {
	if err := validerPreRequisLLM(); err != nil {
		fmt.Printf("\n%s[Erreur Critique] %v%s\n", ColorRed, err, ColorReset)
		return
	}

	purgerMemoireAutomatique()

	c := make(chan os.Signal, 1)
	signal.Notify(c, os.Interrupt, syscall.SIGTERM)
	go func() {
		<-c
		fmt.Println("\n\n" + ColorYellow + "[Arrêt Demandé] Sauvegarde en cours..." + ColorReset)
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
	rapportPath := filepath.Join(desktopPath, "Rapport_Audit_Linguistique.xlsx")
	statePath := filepath.Join(desktopPath, "Audit_State.json")
	stateDB := loadState(statePath)

	var f *excelize.File
	currentRow := 6
	sheet := "Resultats_Audit"

	if _, err := os.Stat(bypassMaxPath(rapportPath)); err == nil {
		fmt.Println(ColorGreen + "[Info] Fichier d'audit existant détecté. Reprise du scan." + ColorReset)
		f, _ = excelize.OpenFile(rapportPath)
		rows, _ := f.GetRows(sheet)
		if len(rows) >= 5 {
			currentRow = len(rows) + 1
		}
	} else {
		f = excelize.NewFile()
		f.SetSheetName("Sheet1", sheet)

		f.SetCellValue(sheet, "A1", "RAPPORT D'AUDIT LINGUISTIQUE (Moteur de Détection Linter)")
		f.SetCellValue(sheet, "A2", "Note : Version archivée. Ce document liste les erreurs isolées par l'IA.")
		f.MergeCell(sheet, "A1", "D1")
		f.MergeCell(sheet, "A2", "D2")

		styleIntro, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Color: "#333333"},
			Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		})
		f.SetCellStyle(sheet, "A1", "A2", styleIntro)

		styleData, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{WrapText: true, Vertical: "top"},
		})
		f.SetColStyle(sheet, "A:D", styleData)

		styleTitre, _ := f.NewStyle(&excelize.Style{
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"#002060"}, Pattern: 1},
			Font:      &excelize.Font{Bold: true, Color: "#FFFFFF"},
			Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center", WrapText: true},
		})

		enTetes := []string{"Coordonnées", "Chemin du fichier (Natif)", "Texte Original (Faute en rouge)", "Suggestion de Correction (En vert)"}
		for i, colName := range enTetes {
			cell := string(rune('A'+i)) + "5"
			f.SetCellValue(sheet, cell, colName)
			f.SetCellStyle(sheet, cell, cell, styleTitre)
		}

		f.SetColWidth(sheet, "A", "A", 30)
		f.SetColWidth(sheet, "B", "B", 40)
		f.SetColWidth(sheet, "C", "D", 70)
	}

	var fichiersOCR []string
	filepath.Walk(bypassMaxPath(dossierCible), func(path string, info os.FileInfo, err error) error {
		if err == nil && !info.IsDir() && strings.HasSuffix(path, ".ocr.txt") {
			fichiersOCR = append(fichiersOCR, path)
		}
		return nil
	})

	if len(fichiersOCR) == 0 {
		fmt.Println(ColorYellow + "Aucun fichier '.ocr.txt' trouvé." + ColorReset)
		return
	}

	var tousLesBlocs []BlockMajeur
	for _, path := range fichiersOCR {
		blocs, err := parserFichierOCR(path)
		if err == nil {
			for _, b := range blocs {
				tousLesBlocs = append(tousLesBlocs, b)
			}
		}
	}

	totalBlocsGlobal := len(tousLesBlocs)
	if totalBlocsGlobal == 0 {
		fmt.Println(ColorYellow + "Traitement annulé : Aucun texte valide." + ColorReset)
		return
	}

	fmt.Printf("%s[Step] Audit en cours (%d micro-blocs) (Ctrl+C pour arrêter)%s\n", ColorBlue, totalBlocsGlobal, ColorReset)

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
			defer func() {
				if r := recover(); r != nil {
					// Ignoré silencieusement pour ne pas polluer la console en mode prod
				}
				<-semaphore
				wg.Done()
			}()

			var batchPhrases []string
			for _, ls := range b.Lignes {
				batchPhrases = append(batchPhrases, ls.Texte)
			}

			if len(batchPhrases) > 0 {
				anomalies := interrogerLLM_Detection(batchPhrases)

				mu.Lock()
				if len(anomalies) > 0 {
					for _, anom := range anomalies {
						for _, ls := range b.Lignes {
							if strings.Contains(ls.Texte, anom.PhraseOriginale) || strings.Contains(anom.PhraseOriginale, ls.Texte) {

								f.SetCellValue(sheet, fmt.Sprintf("A%d", currentRow), ls.Coordonnees)
								f.SetCellValue(sheet, fmt.Sprintf("B%d", currentRow), b.Chemin)

								reReplace := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(anom.MotFaux) + `\b`)
								phraseCorrigee := reReplace.ReplaceAllString(anom.PhraseOriginale, anom.MotCorrige)

								if phraseCorrigee == anom.PhraseOriginale {
									phraseCorrigee = strings.Replace(anom.PhraseOriginale, anom.MotFaux, anom.MotCorrige, 1)
								}

								cibleFaux, cibleCorrige := findExactDiff(anom.PhraseOriginale, phraseCorrigee)

								runsOrig := getRichTextRuns(anom.PhraseOriginale, cibleFaux, "FF0000") // Rouge
								f.SetCellRichText(sheet, fmt.Sprintf("C%d", currentRow), runsOrig)

								runsCorr := getRichTextRuns(phraseCorrigee, cibleCorrige, "00B050") // Vert
								f.SetCellRichText(sheet, fmt.Sprintf("D%d", currentRow), runsCorr)

								currentRow++
								erreursTrouveesGlobal++
								break
							}
						}
					}
				}
				mu.Unlock()
			}

			mu.Lock()
			stateDB.Processed[bKey] = true
			blocsTraitesGlobal++
			afficherBarreProgression(blocsTraitesGlobal, totalBlocsGlobal, "Détection LLM")
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
	fmt.Printf("Fautes réelles isolées et colorisées : %d\n", erreursTrouveesGlobal)
	fmt.Println(ColorGreen + "=========================================" + ColorReset)
}

func main() {
	for {
		atomic.StoreInt32(&stopRequested, 0)
		fmt.Println("\n" + ColorYellow + "=== OUTIL D'AUDIT LINGUISTIQUE (Archive Finale) ===" + ColorReset)
		fmt.Println("1 - Lancer l'analyse (Détection & Colorisation)")
		fmt.Println("2 - Quitter")
		fmt.Print(ColorCyan + "Choix : " + ColorReset)

		scanner := bufio.NewScanner(os.Stdin)
		scanner.Scan()

		switch strings.TrimSpace(scanner.Text()) {
		case "1":
			lancerAuditDetection()
		case "2":
			os.Exit(0)
		default:
			fmt.Println(ColorRed + "Saisie invalide." + ColorReset)
		}
	}
}
