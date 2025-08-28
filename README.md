# go-office-documents-finder
This project was made to find office documents (zip archives with embedded xml) correct extensions after being recovered as zips. And also working on an llm implementation to find a coherent name for pptx and xlsx documents.
So, in the first place this was made to recover the file names of zip archives and the files extensions.
The actual implementation isn't good enough to have some coherent names (5 most common words in the extracted file text).
Further version will potentially have a filtering system to sort files based on the content.
It can search through folders recursively.

to use it:
- change the path to your drive/folder
- specify the count of most common words to make the title (default 5), in the function count recurrence (change n variable)
- coming soon : launch ollama (llama3)


potentially comming soon:
- some llm to handle filenames
- some sorting algorithm (ollama self hosted)
- xlsx files handling


How dooes this work?
this programm checks the occurence of bytes in the zip architecture, it is possible because of the structure of office files. Office files are made with a folder conaining the data (the text), called after the correct extension, therefore searching in the bytes of the file is not a bad idea. Also found out that the docx files name is mostly intact in the metadata, so if we search in docProps/core.xml (were the metadata is stored) we can get the title. The problem is that it's only valid for docx and not for pptx and xlxs (not working with other office extensions for now). Not having any way of backing the original name up, you still gotta make an understandable name for the user, this is why there is an implementation of the five most common words. This implementation is still better than hashed file names but it's not really connected, words don't make sense. To pally that, an implementation of a basic llm is being coded (ollama api).
