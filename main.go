package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
)

func scandirs(path string) ([]string, []string) {
	var outputdir []string
	var outputformat []string

	var walkn func(p string)
	walkn = func(p string) {
		entries, err := os.ReadDir(p)
		if err != nil {
			log.Printf("Failed to read directory %s: %v", p, err)
			return
		}
		for _, entry := range entries {
			fullPath := filepath.Join(p, entry.Name())
			if entry.IsDir() {
				outputdir = append(outputdir, fullPath)
				walkn(fullPath)
			} else {
				ext := filepath.Ext(entry.Name())
				switch ext {
				case ".zip", ".docx", ".pptx", ".xlsx":
					outputformat = append(outputformat, fullPath)
				}
			}
		}
	}

	walkn(path)
	return outputdir, outputformat
}

func ident(path string) (result string) {
	filext := filepath.Ext(path)
	cleaned := strings.TrimSuffix(path, filext)
	file, err := os.Open(path)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	data, err := io.ReadAll(file)
	if checkpres(data, "word/") {
		return cleaned + ".docx"
	} else if checkpres(data, "xl/") {
		return cleaned + ".xlxs"
	} else if checkpres(data, "ppt/") {
		return cleaned + ".pptx"
	}

	return path
}
func checkpres(data []byte, rech string) (found bool) {
	found = bytes.Contains(data, []byte(rech)) //mind to maybe change to bytes.Index if it doesnt work and ad >0 to return bool
	return found
}
func readtitle(path string) string {
	file, err := zip.OpenReader(path)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	var docx1 io.ReadCloser
	for _, f := range file.File {
		if f.Name == "docProps/core.xml" {
			docx1, err = f.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer docx1.Close()
			break
		}
	}
	if docx1 == nil {
		log.Fatal("doc doesnt exist")
	}
	xmldecode := xml.NewDecoder(docx1)
	var extr strings.Builder
	var final string
	for {
		token, _ := xmldecode.Token()
		if token == nil {
			break
		}
		switch tok := token.(type) {
		case xml.StartElement:
			if tok.Name.Local == "title" {
				var inter string
				xmldecode.DecodeElement(&inter, &tok)
				extr.WriteString(inter)
			}

		}
		final = strings.ReplaceAll(extr.String(), " ", "_")
	}
	return final
}

func tokenization(text string) []string {
	var alphanum = regexp.MustCompile(`[^\p{L}\p{N} ]+`)
	ntext := alphanum.ReplaceAllString(text, "")
	spacer := regexp.MustCompile(`\s+`)
	d := spacer.ReplaceAllLiteralString(ntext, " ")
	d = strings.TrimSpace(d)
	splittedstr := strings.Split(d, " ")
	return splittedstr
}

func checkstopword(stopwords []string, word string) bool {
	for _, w := range stopwords {
		if w == word {
			return true
		}
	}
	return false
}

func removeStopWords(text []string) []string {
	var stopwords = []string{
		"a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are", "aren't", "as", "at",
		"be", "because", "been", "before", "being", "below", "between", "both", "but", "by",
		"could", "couldn't",
		"did", "didn't", "do", "does", "doesn't", "doing", "don't", "down", "during",
		"each",
		"few", "for", "from", "further",
		"had", "hadn't", "has", "hasn't", "have", "haven't", "having", "he", "he'd", "he'll", "he's", "her", "here", "here's", "hers", "herself", "him", "himself", "his", "how", "how's",
		"i", "i'd", "i'll", "i'm", "i've", "if", "in", "into", "is", "isn't", "it", "it's", "its", "itself",
		"let's",
		"me", "more", "most", "mustn't", "my", "myself",
		"no", "nor", "not", "now",
		"of", "off", "on", "once", "only", "or", "other", "ought", "our", "ours", "ourselves", "out", "over", "own",
		"same", "she", "she'd", "she'll", "she's", "should", "shouldn't", "so", "some", "such",
		"than", "that", "that's", "the", "their", "theirs", "them", "themselves", "then", "there", "there's", "these", "they", "they'd", "they'll", "they're", "they've", "this", "those", "through", "to", "too",
		"under", "until", "up",
		"very",
		"was", "wasn't", "we", "we'd", "we'll", "we're", "we've", "were", "weren't", "what", "what's", "when", "when's", "where", "where's", "which", "while", "who", "who's", "whom", "why", "why's", "will", "with", "won't", "would", "wouldn't",
		"you", "you'd", "you'll", "you're", "you've", "your", "yours", "yourself", "yourselves",

		"alors", "au", "aucuns", "aussi", "autre", "avant", "avec", "avoir",
		"bon",
		"car", "ce", "cela", "ces", "ceux", "chez", "comme", "comment",
		"dans", "des", "du", "donc", "dos",
		"elle", "elles", "en", "encore", "est", "et", "eu", "être",
		"fait", "faites", "fois",
		"hommes",
		"ici", "il", "ils",
		"je", "juste",
		"la", "le", "les", "leur", "là",
		"ma", "maintenant", "mais", "me", "même", "mes", "mien", "moins", "mon", "mot",
		"ni", "nom", "nos", "notre", "nous",
		"on", "ou", "où",
		"par", "parce", "pas", "peu", "peut", "pour", "pourquoi", "quand", "que", "quel", "quelle", "quelles", "quels", "qui",
		"sa", "sans", "ses", "seulement", "si", "sien", "son", "sont", "sous", "soyez", "sujet", "sur",
		"ta", "te", "tes", "toi", "ton", "tous", "tout", "trop", "très", "tu",
		"un", "une",
		"vos", "votre", "vous",
		"vu",
		"ça", "étaient", "état", "était", "étant", "être",

		"aber", "alle", "allem", "allen", "aller", "alles", "als", "also", "am", "an", "ander", "andere", "anderem", "anderen", "anderer", "anderes",
		"anderm", "andern", "anderr", "anders",
		"auch", "auf", "aus",
		"bei", "bin", "bis", "bist",
		"da", "damit", "dann", "das", "dass", "dasselbe", "dazu", "dein", "deine", "deinem", "deinen", "deiner", "deines", "dem", "den", "der", "derer", "dessen", "des", "deshalb", "die", "dies", "diese", "diesem", "diesen", "dieser", "dieses",
		"doch", "dort", "du", "durch",
		"ein", "eine", "einem", "einen", "einer", "eines", "einig", "einige", "einigem", "einigen", "einiger", "einiges", "einmal", "er", "es",
		"etwas",
		"für",
		"gegen", "gewesen",
		"hab", "habe", "haben", "hat", "hatte", "hatten", "hier", "hin", "hinter",
		"ich", "ihm", "ihn", "ihnen", "ihr", "ihre", "ihrem", "ihren", "ihrer", "ihres", "im", "in", "indem", "ins", "ist",
		"jede", "jedem", "jeden", "jeder", "jedes",
		"jener", "jenes",
		"jetzt",
		"kann", "kein", "keine", "keinem", "keinen", "keiner", "keines", "könnte",
		"machen", "man", "manche", "manchem", "manchen", "mancher", "manches", "mein", "meine", "meinem", "meinen", "meiner", "meines", "mich", "mir", "mit", "muss", "musste",
		"nach", "nicht", "nichts", "noch", "nun", "nur",
		"ob", "oder", "ohne",
		"sehr", "sein", "seine", "seinem", "seinen", "seiner", "seines", "selbst", "sich", "sie", "sind", "so", "solche", "solchem", "solchen", "solcher", "solches", "soll", "sollte", "sondern", "sonst",
		"über", "um", "und", "uns", "unser", "unsere", "unserem", "unseren", "unserer", "unseres",
		"unter",
		"viel", "vom", "von", "vor",
		"war", "waren", "warst", "was", "weg", "weil", "weiter", "welche", "welchem", "welchen", "welcher", "welches", "wenn", "werde", "werden", "wie", "wieder", "will", "wir", "wird", "wirst", "wo", "wollen", "wollte", "worden", "wurde", "wurden",
		"zu", "zum", "zur", "zwar", "zwischen",
	}
	filtered := []string{}
	for _, word := range text {
		if !checkstopword(stopwords, word) {
			filtered = append(filtered, word)
		}
	}
	return filtered
}

type Slide struct {
	Name string
	Obj  io.ReadCloser
}

func readpdf(path string) string {
	pptx, err := zip.OpenReader(path)
	if err != nil {
		log.Fatal(err)
	}
	defer pptx.Close()
	var extracted strings.Builder
	var slides []Slide
	for _, f := range pptx.File {
		if strings.HasPrefix(f.Name, "ppt/slides/") && strings.HasSuffix(f.Name, ".xml") {
			rder, err := f.Open()
			if err != nil {
				log.Printf("error file", f.Name, err)
				continue
			}
			defer rder.Close()
			slides = append(slides, Slide{
				Name: f.Name,
				Obj:  rder,
			})
		}
	}
	for _, slide := range slides {
		defer slide.Obj.Close()
		decoder := xml.NewDecoder(slide.Obj)
		for {
			token, err := decoder.Token()
			if err != nil {
				break
			}
			switch to := token.(type) {
			case xml.StartElement:
				if to.Name.Local == "t" {
					var inter string
					err := decoder.DecodeElement(&inter, &to)
					if err == nil {
						extracted.WriteString(inter + "")
					}
				}
			}
		}
	}
	return extracted.String()
}

func countRecurence(tokenizedtext []string) []string {
	var mostpresent []string
	count := make(map[string]int)
	type orderedslice struct {
		word  string
		count int
	}
	var slice []orderedslice

	for _, word := range tokenizedtext {
		count[word]++

	}

	for key, value := range count {
		slice = append(slice, orderedslice{key, value})
	}
	sort.Slice(slice, func(i, j int) bool {
		return slice[i].count > slice[j].count
	})
	n := 5
	if len(slice) < 5 {
		n = len(slice)
	}

	for i := 0; i < n; i++ {
		mostpresent = append(mostpresent, slice[i].word)
	}
	return mostpresent
}

func dFName(mostpresent []string, fext string) string {
	fname := strings.Join(mostpresent, "_")
	return fname + fext
}

func main() {
	path := "path/to/your/files" //change that later on ***********************************************************
	dirs, files := scandirs(path)
	log.Printf("\nfiles:\n", files)
	var fileschanged []string
	for _, f := range files {
		newext := ident(f)
		log.Printf("\n\nnewextension:", newext)
		err := os.Rename(f, newext)
		if err != nil {
			log.Printf("couldnt change that file: ", newext, err)
			continue
		}
		fmt.Print("successfully changed the extension \n")
		fileschanged = append(fileschanged, newext)
	}
	fmt.Printf("files changed \n: ", fileschanged)

	for _, dir := range dirs {
		fmt.Print("\n\ndirectories handled.........................................\n\n\n", dir, "\n\n")
	}

	for _, fi := range fileschanged {
		subpath := filepath.Dir(fi) + "/"
		fmt.Print("\nfilepath........ ", fi, " ..........\n")
		fmt.Printf("\nLaunching renaming procedure...\n...\n")
		filext := filepath.Ext(fi)
		if filext == ".docx" {
			newname := readtitle(fi)
			tk := tokenization(newname)
			var name string
			for _, t := range tk {
				fmt.Print(t, "\n")
				if t != "/" {
					name += t
				}
			}
			err := os.Rename(fi, subpath+"/"+name+filext)
			if err != nil {
				log.Printf("\nfailed to retrieve docx name: ", fi, "error: ", err)
			}

		} else if filext == ".pptx" {
			fmt.Printf("\nrenaming a ppt.....\n")
			text := readpdf(fi)
			tk := tokenization(text)
			cleaned := removeStopWords(tk)
			mostpresent := countRecurence(cleaned)
			finalname := dFName(mostpresent, filext)
			fmt.Print("\nfinalname: ", finalname, "\n", "path: \n", subpath, "\n............................................................\n")
			err := os.Rename(fi, filepath.Join(subpath, finalname))
			if err != nil {
				log.Printf("\nfailed to retrieve pptx:", fi, "error", err)
			}
		} else if filext == ".xlsx" {
			continue
		}

	}
}
