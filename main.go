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

func removeStopWords(text []string, stopwords []string) []string {
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
	fpath := "test.txt"
	data, err := os.ReadFile(fpath)
	if err != nil {
		log.Fatal("couldnt read stopwords list please check if its in ur working directory")
	}
	stopwords := strings.Split(string(data), "\n")
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
			cleaned := removeStopWords(tk, stopwords)
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
