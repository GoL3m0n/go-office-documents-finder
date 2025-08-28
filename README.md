# go-office-documents-finder
This project was made to find office documents (zip archives with embedded xml) correct extensions after being recovered as zips. And also working on an llm implementation to find a coherent name for pptx and xlsx documents.
So, in the first place this was made to recover the file names of zip archives and the files extensions.
The actual implementation isn't good enough to have some coherent names (5 most common words in the extracted file text).
Further version will potentially have a filtering system to sort files based on the content.
It can search through folders recursively.

to use it:
- change the path to your drive/folder
- specify the count of most common words to make the title (default 5), in the function count recurrence (change n variable)


potentially comming soon:
- some llm to handle filenames
- some sorting algorithm (ollama self hosted)
- xlsx files handling
