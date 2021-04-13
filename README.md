# PdfConverter

Simple application to convert .doc, .docx, .ppt, .pptx files to .pdf format.

**_Only available for Windows_**

## Instructions

Install with [pip](https://pip.pypa.io/en/stable/) all modules in *requirements.txt* file. Run with *python* converter.py, with documents as arguments. The pdf files will be created in the same directory.

## Tips

I added file to sendto menu, for easier use. To do that, add file to shell:sendto folder.

I used this batch file in sendto folder.

```bash
python.exe converter.py %*
```

This way you can select files, and convert them in place with just few clicks :D
