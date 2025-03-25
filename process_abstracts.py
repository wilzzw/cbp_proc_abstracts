##################### Created by Wilson Zeng on April 27th, 2019 ##############################

# Dependencies (required softwares and packages: 1. BeautifulSoup, 2. Pandoc

from bs4 import BeautifulSoup
import subprocess
import re
import os

##################### Mutable Inputs ######################

WORKING_DIR = "." # The directory where the script is run
# The outputs will be written to OUTPUT_TEX
OUTPUT_TEX = 'output.tex'


# !Potential manual work needed before executing the script: fixing typos in submitted filenames
# !Sometimes people misspell their own names.
# !Note: if the abstract file format is doc instead of docx, this code will not work. People should not be using doc these days. Accommodate and process manually if needed.
# Problematic abstracts will be printed out upon execution of the script. They require to be manually checked to see what is wrong and manually processed.

# Available image extensions
# TODO: distinguish pdf image and some pdf abstract text
IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.tiff', '.pdf']

# NOTE: As far as I am aware, PANDOC processes '&' and '%' symbols correctly.
# Do not double-process them by adding to the dictionary special_characters!!
# TODO: consider using a package to handle unicode to latex conversion.
# Greek letters
GREEK = {'α': r'$\alpha$', 'β': r'$\beta$', 'γ': r'$\gamma$', 'δ': r'$\delta$',
         'ε': r'$\epsilon$', 'ζ': r'$\zeta$', 'η': r'$\eta$', 'θ': r'$\theta$',
         'ι': r'$\iota$', 'κ': r'$\kappa$', 'λ': r'$\lambda$', 'µ': r'$\mu$',
         'ν': r'$\nu$', 'ξ': r'$\xi$', 'π': r'$\pi$', 'ρ': r'$\rho$', 'σ': r'$\sigma$',
         'τ': r'$\tau$', 'υ': r'$\upsilon$', 'φ': r'$\phi$', 'χ': r'$\chi$', 'ψ': r'$\psi$', 'ω': r'$\omega$'}
# Latin with accents
LATIN_WITH_ACCENTS = {'é': r'\'{e}', 'è': r'\`{e}', 'ü': r'\"{u}', 'ä': r'\"{a}', 'ö': r'\"{o}'}

# The symbol I know PANDOC does not handle properly
# PANDOC gives '\textasciitilde{}' for tilde, but it ends up being the upper tilde symbol in latex
# So, we need to replace it with the middle tilde symbol in latex "\sim"
# PANDOC handles most special characters already, except "~", which is addressed in PANDOC_FIX (2019-04-28)
# If there are any other special characters that PANDOC does not parse correctly, the user might want to add it to PANDOC_FIX dictionary.
PANDOC_FIX = {
                r'\textasciitilde{}': '$\sim$'
             }

# Pool all fixes together. These fixes will be applied after we have allowed PANDOC to do its job (i.e. To clean up what PANDOC did not do a great job in).
SYMBOL_FIXES = {**GREEK, **LATIN_WITH_ACCENTS, **PANDOC_FIX} 

# Name segments that should not be capitalized
# Anything else? Please update :)
DO_NOT_CAPITALIZE = ['van', 'der', 'van\'t']


##################### Get working directory contents ######################
# Extract file names in the current working directory into tuple: (file_name, extension)
# Only extract files with extension .docx and image_ext

all_files = [os.path.splitext(f) for f in os.listdir(WORKING_DIR) if os.path.isfile(f)]
doc_files = [(filename, extension) for filename, extension in all_files if extension == '.docx']

# This code will find image files associated with the following extensions in the current
# directory by matching the image file name against Lastname_Firstname.docx
fig_files = [(filename, extension) for filename, extension in all_files if extension in IMAGE_EXTENSIONS]

# Sort by submitters' last names alphabetically
doc_files = sorted(doc_files, key=lambda x: x[0])

# Clear current output file if available
latex = open(OUTPUT_TEX, 'w')
latex.close()

################################## Functions ######################################
# General conversion of text extracted from html into latex format
# First let PANDOC handles the dirty work
# Then apply our specific fixes
# text: str, the text to be converted
# Returns: str, the input text in latex format
def textfix(text):
    if '\n' in text:
        paragraphs = text.split('\n')
    else:
        paragraphs = [text]

    fixed_paragraphs = []
    for p in paragraphs:
        # Call PANDOC to convert the html text to latex
        s = subprocess.run(['pandoc', '-f', 'html', '-t', 'latex'], input=p, stdout=subprocess.PIPE, universal_newlines=True)
        # The output of PANDOC is the converted format of the text
        fixed_text = s.stdout[:-1]

        # PANDOC does not handle the symbols in SYMBOL_FIXES correctly
        # Apply fixes to the symbols in the text that needs to be converted to latex format
        for fix in SYMBOL_FIXES.keys():
            fixed_text = fixed_text.replace(fix, SYMBOL_FIXES[fix])

        # PANDOC has a side effect: it strips off the leading/trailing white spaces
        # Add them back to the text
        fixed_text = ' '*(len(p)-len(p.lstrip())) + fixed_text + ' '*(len(p)-len(p.rstrip()))
        # Not sure what the following line is for anymore...
        fixed_text = fixed_text.replace('\n', ' ')
        # Append the fixed_text to the list of fixed_paragraphs
        fixed_paragraphs.append(fixed_text)
    # Join the fixed_paragraphs into a single string with latex newline characters
    total_fixed_text = r'\\'.join(fixed_paragraphs)
    return total_fixed_text

# Function to be called to wrrap the abstract content in commands
def choose_template(figure=False):
    if not figure:
        return r'\posterAbstractSansFigure'
    return r'\posterAbstractWithFigure'


def brace_text(text):
    return f"{{{text}}}"

# Function to be called to process abstract title in latex
# i.e. Place it in brackets
def proc_title(title=''):
    return brace_text(title)

# Function to be called to parse an author's information (variable 'author')
# An author information is a tuple: (presenting: bool, First_name, ..., Last_name)
# If the author is the presenting author, presenting == True, and vice versa
# Returns the name string of the author separated by spaces, as well as whether the author is presenting
def parse_author_info(author: tuple):
    presenting = author[0]
    name_components = []
    for name in author[1:]:
        if name in DO_NOT_CAPITALIZE:
            name_components.append(textfix(name))
        else:
            name_components.append(textfix(name.capitalize()))
    if presenting:
        return {'name': r'\textbf{'+' '.join(name_components)+'}', 'presenting?': presenting}
    return {'name': ' '.join(name_components), 'presenting?': presenting}

# # Function to be called to process a list of authors' information and the corresponding affiliation numbers
# # authors are a list of author information tuples; affiliations are a list of lists of corresponding affiliation numbers (as strings of integers)
def proc_authors(authors, affiliations):
    author_list = []
    label = ''
    for author, assoc_affil in zip(authors, affiliations):
        author_info = parse_author_info(author)
        author_name = author_info['name']
        # # Is the author the presenting author?
        # if author_info['presenting?']:
        #     author_name = r'\textbf{'+author_name+'}'
        # Make the latex string for the author, with superscripted affiliation numbers
        author_string = author_name+',$^{'+','.join(assoc_affil)+'}$'
        author_list.append(author_string)
    # Join the author_list into a single string in latex format
    author_tex_string = ' '.join(author_list)
    return brace_text(author_tex_string)

# Function to be called to produce affiliations in latex format
def proc_affiliations(affil_info: dict):
    affiliation_texts = ['{']
    # n is the affiliation number
    for n in sorted(list(affil_info.keys())):
        # Last affiliation does not need a newline symbol at the end
        if n == max(list(affil_info.keys())):
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n]))
        else:
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n])+r'\\ ')
    affiliation_texts.append('}')
    return "".join(affiliation_texts)

# Function to be called to write poster number in latex format
def proc_poster_number(poster_number):
    # return r'{P\#}'
    return brace_text(poster_number)

# # Function to be called to write lines to insert figure in latex format
def proc_figure(fig_file):
    return brace_text(fig_file)

# Function to be called to produce references in latex format
def proc_ref(ref_info):
    reference_texts = ['{']
    for n in sorted(list(ref_info.keys())):
        if n == max(list(ref_info.keys())):
            reference_texts.append('{['+str(n)+']} '+ref_info[n])
        else:
            reference_texts.append('{['+str(n)+']} '+ref_info[n]+r'\\')
    reference_texts.append('}')
    return reference_texts

def write_abstract_latex(output_tex, abstract_title, authors, associated_affiliations, affiliations_list, abstract_text, figure_file, poster_number=0):
    fig_avail = len(figure_file) > 0
    with open(output_tex, 'a') as f:
        # Choose template
        f.write(choose_template(figure=fig_avail))
        f.write(proc_poster_number(poster_number))
        f.write(proc_title(abstract_title))
        f.write(proc_authors(authors, associated_affiliations))
        f.write(proc_affiliations(affiliations_list))
        if fig_avail:
            f.write(proc_figure(figure_file))
        f.write(brace_text(abstract_text))
        f.write("\n")
    return

##################### Main Script ######################

# for file_info in doc_files:
for docname, extension in doc_files:
    ##################### Part I. Parsing Word File into HTML with PANDOC #####################
    abstract_doc = docname + extension
    print('Processing '+abstract_doc)
    html_name = f'{docname}.html'
    # Call subprocess to run: pandoc $abstract_doc -f docx -t html -o $html_name
    subprocess.run(['pandoc', abstract_doc, '-f', 'docx', '-t', 'html', '-o', html_name])

    ######################### Part II. Look for a matching figure file #########################
    # Look whether submitter submitted a figure file
    # Matching the names of the figure files with the docx file
    matching_figures = [figname + ext for figname, ext in fig_files if figname.lower() == docname.lower()]

    if len(matching_figures) > 0:
        fig_name = matching_figures[0]
    else:
        fig_name = ''
    if len(fig_name) > 0:
        print('Submitter has a figure: '+fig_name+'\n')

    ############ Part III. Extracting Contents from HTML File using BeautifulSoup ###############
    # Parse html file into BeautifulSoup :) to extracting the contents
    # See https://www.crummy.com/software/BeautifulSoup/bs4/doc/ for Beautiful Soup docs :)
    f = open(html_name, 'r')
    soup = BeautifulSoup(f, 'html.parser')

    # Extract all table/cell contents into a list of table objects
    # The abstract doc file consists of tables after all; they are transfered into html
    all_tables = soup.find_all('table')

    # If someone messed up the document so badly/did not use the submission template
    # i.e. no tables found
    if len(all_tables) == 0:
        print(abstract_doc+' is problematic. No tables were found. Please check the docx file manually as the submitter likely did not use the provided abstract template.\n')
        continue

    # The first table has the abstract title
    # abstract_title_components = [textfix(str(c)) for c in all_tables[0].find_all(['th', 'td'])[0].text]
    abstract_title_components = [textfix(str(c)) for c in all_tables[0].find_all(['th', 'td'])[0].contents]
    abstract_title = ''.join(abstract_title_components) ### Processed abstract title in latex format!

    # The second table has the the authors & affiliations
    # Extract all fields ('tr') in the authors & affiliations table
    author_and_affil_fields = all_tables[1].find_all('tr')
    author_list = []
    associated_affiliations = []
    for entry in author_and_affil_fields[1:]:
        # print(entry)
        # Address the issue where an empty row or non author row occurs (without number or text) which results in ValueError
        num_author = re.sub('[^0-9]', '', entry.find_all('td')[0].text)
        if len(num_author) == 0:
            continue
        num_author = int(num_author)
        author_name = entry.find_all('td')[1].text
        presenting = "*" in author_name
        if presenting:
            author_name = author_name.replace('*', '')
        # If no author in this line, skip over
        if len(author_name) == 0:
            continue
        author_name_split = author_name.split(' ')
        author_affil = [a for a in re.sub('[^0-9]', ' ', entry.find_all('td')[2].text).split(' ') if a != '']
        author_info = (presenting,) + tuple(author_name_split)
        author_list.append(author_info)
        associated_affiliations.append(author_affil)
    # The format of the author_list and associated_affil: author informations are tuples: tuple(presenting?, from first name to last name); affiliations are lists of corresponding affiliation numbers as strings (not integers).

    # The third table has the affiliation names
    # Extract all fields ('tr') in the affiliations
    affiliation_fields = all_tables[2].find_all('tr')
    affiliation_dict = {}
    for entry in affiliation_fields:
        # print(entry)
        # Address the issue where an empty row occurs (without number or text) which results in ValueError
        affil_num = re.sub('[^0-9]', '', entry.find_all(['td', 'th'])[0].text)
        if len(affil_num) == 0:
            continue
        affil_num = int(affil_num)
        affil = entry.find_all(['td', 'th'])[1]
        if affil.p != None:
            affil = affil.p.text
        else:
            affil = affil.text
        if len(affil) > 0:
            affiliation_dict[affil_num] = affil

    # The fourth table has the abstract body text
    abstract_content = [textfix(str(c)) for c in all_tables[3].find_all(['th', 'td'])[0].contents]
    abstract_main = ''.join(abstract_content) ### Processed abstract text in latex format!

    # The fifth table has the references
    # Extract all fields ('tr') in the references

    # if len(all_tables) > 4:
    #     reference_fields = all_tables[4].find_all('tr')
    #     ref_list = {}
    #     for entry in reference_fields:
    #         ref_num = int(re.sub('[^0-9]', '', entry.find_all(['td', 'th'])[0].text))
    #         ref = entry.find_all(['td', 'th'])[1]
    #         fancy_texts = [str(x) for x in ref.find_all(['em', 'strong', 'sup'])]
    #         raw_ref_text = entry.find_all(['td', 'th'])[1]
    #         raw_ref_text = ''.join([textfix(str(c)) for c in raw_ref_text.contents])
    #         if len(raw_ref_text) > 0:
    #             ref_list[ref_num] = raw_ref_text

    f.close()

    ##################### Part II. Writing TEX File #####################
    write_abstract_latex(OUTPUT_TEX, abstract_title, author_list, associated_affiliations, affiliation_dict, abstract_main, fig_name)
    print('\n')

# Remove temporary/intermediate html files
html_files = [f for f in os.listdir('.') if os.path.isfile(f) and os.path.splitext(f)[1] == '.html']
for htm in html_files:
    os.remove(htm)