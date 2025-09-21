package main
/*this verrsion is shorter but contains a version with an llm using the ollama api
feel free to change to a lighter version, for a shorter execution time and less ressource consumption*/



import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
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

type request struct {
	Model  string `json:"model"`
	Prompt string `json:"prompt"`
	Stream bool   `json:"stream"`
}
type response struct {
	Response string `json:"response"`
}

// working on the llm, will maybe need to prompt engineer but actually it is fine.
func llmname(fulltext string, fext string) string {
	url := "http://localhost:11434/api/generate"
	prompt := "find file name for extracted text of " + fext + " : " + fulltext + "\n max 5 words,separate each with underscores,dont add verbose,return filename"
	req := request{
		Model:  "llama3",
		Prompt: prompt,
		Stream: false,
	}

	jsonf, err := json.Marshal(req)
	if err != nil {
		log.Printf("error crafting json")
	}
	resp, err := http.Post(url, "application/json", bytes.NewBuffer(jsonf))
	if err != nil {
		log.Printf("couldn't send request")
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		log.Printf("couldnt read response")
	}
	var responsef response
	err = json.Unmarshal(body, &responsef)
	if err != nil {
		log.Print(err)
	}
	return responsef.Response

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
			finalname := llmname(text, filext)
			err := os.Rename(fi, filepath.Join(subpath, finalname))
			if err != nil {
				log.Printf("\nfailed to retrieve pptx:", fi, "error", err)
			}
		} else if filext == ".xlsx" {
			continue
		}

	}
}
