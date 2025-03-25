# Contact
wilsonzzw@gmail.com

## What is it?
This is a script I created to process CBP abstracts into latex format.

## How does it work?

The script reads the docx file and convert into a more parseable html format using PANDOC.
Then, it uses BeautifulSoup python library to parse the html and extract the abstract contents (i.e. title, authors and affiliations, main body).

## Requirements

- Python 3
- Pandoc (https://pandoc.org/installing.html); working as of version 2.5-3build2
- BeautifulSoup4 (https://beautiful-soup-4.readthedocs.io/en/latest/#installing-beautiful-soup); working as of version 4.11.1

## To run the script

- First, place all submitted abstract files in the same directory as the script.
- Then, run the script using the following command:
```
python process_abstracts.py
```