
# Description
This is a script that converts PDF file into images, then into a PowerPoint .pptx file.

# Install

## Step 1: Install Poppler on mac, a pdf rendering library

Follow [this instruction](http://macappstore.org/poppler/) if your Mac has not installed Homebrew yet.

Run this command in terminal:

```shell
brew install poppler
```

## Step 2: install python libraries needed
Open terminal and cd into the folder `pptx_converter`, then run:

```shell
pip3 install -r requirements.txt
```

# Usage
Open terminal and cd into the folder `pptx_converter`, then run:


```shell
python3 pdf2pptx.py --from <path_to_pdf_file> --to <path_to_pptx_file> 
```

```
optional arguments:
  -h, --help            show this help message and exit
  -f PDF_PATH, --from PDF_PATH
                        Pdf file path (REQUIRED)
  -t PPTX_PATH, --to PPTX_PATH
                        pptx file path (REQUIRED)
  -d DPI, --dpi DPI     DPI for the jpg images converted from PDF, defaulted to 600
  -v, --verbose         increase output verbosity
```

Example:

```shell
python3 pdf2pptx.py --from ./test.pdf --to ./test.pptx --dpi 1200
```

