## vbimg2pdf
Convert jpeg/png images to multi-page pdf file

### Description

Can be used to embed jpeg/png images in a single pdf file without resizing and recompressing input format. Uses `Microsoft Print to PDF` printer as pdf writer.

### Sample usage

 - Generate `output.pdf` from all jpegs in current folder
```
c:> vbimg2pdf.exe *.jpg -o output.pdf
```
### Command-line
```
vbimg2pdf 0.1 (c) 2018 by wqweto@gmail.com
Convert jpeg/png images to multi-page pdf file

Usage: vbimg2pdf.exe [options] <in_file.jpg> ...

Options:
  -o OUTFILE         write result to OUTFILE
  -paper SIZE        output paper size (e.g. A4)
  -orientation ORNT  page orientation (e.g. portrait)
  -margins L[/T/R/B] page margins in inches (e.g. 0.25)
  -q                 in quiet operation outputs only errors
  -nologo            suppress startup banner
```
