'''
This program is designed to convert google forms submitted answers to a latex book of abstracts.


Petr Kurka

UFE
20.06.2023
'''


import tkinter as tk
import time
import copy
from tkinter import filedialog
import unicodedata
import csv
import re
import os

NUMBERE_OF_UNSUPORTED_CHAR = 0
NUMBERE_OF_UNSUPORTED_GREEK = 0

# Define the ASCII art for the letters
u = '''
 _   _
| | | |
| | | |
| |_| |
 \___/
'''

f = '''
 ______
|  ____|
| |__
|  __|
|_|
'''

e = """
 _____
|  ___|
| |_
|  _|
| |___
|_____|
"""

# Print each letter with a slight delay in between
for letter in [u, f, e]:
    print(letter)
    time.sleep(0.3)

print('Petr Kurka 2023\n\n\n')
time.sleep(0.5)

# Is letter in extended ascii?
def is_extended_ascii(char):
    return ord(char) >= 128 and ord(char) <= 255

# Define a list of characters in the Czech alphabet
czech_alphabet = [
    'a', 'á', 'b', 'c', 'č', 'd', 'ď', 'e', 'é', 'ě', 'f', 'g', 'h', 'ch', 'i', 'í', 'j',
    'k', 'l', 'm', 'n', 'ň', 'o', 'ó', 'p', 'q', 'r', 'ř', 's', 'š', 't', 'ť', 'u', 'ú', 'ů',
    'v', 'w', 'x', 'y', 'ý', 'z', 'ž'
]

# Define a dictionary of LaTeX Greek letter replacements
def replace_greek_letters(s):
    greek_letters = {
        'α': r'$\alpha$',
        'β': r'$\beta$',
        'γ': r'$\gamma$',
        'δ': r'$\delta$',
        'ε': r'$\epsilon$',
        'ζ': r'$\zeta$',
        'η': r'$\eta$',
        'θ': r'$\theta$',
        'ι': r'$\iota$',
        'κ': r'$\kappa$',
        'λ': r'$\lambda$',
        'μ': r'$\mu$',
        'ν': r'$\nu$',
        'ξ': r'$\xi$',
        'ο': r'$\omicron$',
        'π': r'$\pi$',
        'ρ': r'$\rho$',
        'σ': r'$\sigma$',
        'τ': r'$\tau$',
        'υ': r'$\upsilon$',
        'φ': r'$\phi$',
        'χ': r'$\chi$',
        'ψ': r'$\psi$',
        'ω': r'$\omega$',
        'Α': r'$A$',
        'Β': r'$B$',
        'Γ': r'$\Gamma$',
        'Δ': r'$\Delta$',
        'Ε': r'$E$',
        'Ζ': r'$Z$',
        'Η': r'$H$',
        'Θ': r'$\Theta$',
        'Ι': r'$I$',
        'Κ': r'$K$',
        'Λ': r'$\Lambda$',
        'Μ': r'$M$',
        'Ν': r'$N$',
        'Ξ': r'$\Xi$',
        'Ο': r'$O$',
        'Π': r'$\Pi$',
        'Ρ': r'$P$',
        'Σ': r'$\Sigma$',
        'Τ': r'$T$',
        'Υ': r'$\Upsilon$',
        'Φ': r'$\Phi$',
        'Χ': r'$X$',
        'Ψ': r'$\Psi$',
        'Ω': r'$\Omega$',
    }
    # Use a regular expression to find all Greek letters
    greek_pattern = re.compile('[α-ωΑ-Ω]')
    # Replace each Greek letter with its LaTeX equivalent in $ symbols
    return greek_pattern.sub(lambda x: greek_letters[x.group()], s)

# Define a function to check if a character is GREEK
def is_greek_letter(char):
    category = unicodedata.category(char)
    return category in ('Ll', 'Lu') and unicodedata.name(char).startswith('GREEK')

# Define a function to check if a character is in the Czech alphabet
def is_czech_char(char):
    return char.lower() in czech_alphabet

# Replace the special characters with their LaTeX escaped versions
def escape_latex(text):
    global NUMBERE_OF_UNSUPORTED_CHAR
    global NUMBERE_OF_UNSUPORTED_GREEK
    escaped_text = ''
    for char in text:
        if is_greek_letter(char):
            escaped_text += replace_greek_letters(char)
            NUMBERE_OF_UNSUPORTED_GREEK += 1
        # Check if the character is a Unicode character
        elif (not char.isascii() and not is_extended_ascii(char) and not is_czech_char(char) and not(char in "‐-–'—")) or char in '*_{}^%&#+\$':
            # Replace the character with the string '|UNICODE|'
            escaped_text += r'\colorbox{Red}{\#}'
            NUMBERE_OF_UNSUPORTED_CHAR += 1
        else:
            # Otherwise, add the character to the new string
            escaped_text += char
    return escaped_text

# Generates LATEX file
def GenerateStructureFile(colour):
    A = r'''
    %----------------------------------------------------------------------------------------
    %	PACKAGES AND OTHER DOCUMENT CONFIGURATIONS
    %----------------------------------------------------------------------------------------

    \usepackage[czech]{babel} % základní podpora pro češtinu, mj. správné dělení slov
    \usepackage[utf8]{inputenc} % vstupní kódování je UTF-8
    \usepackage[T1]{fontenc} % výstupní kódování

    \usepackage{microtype} % Better typography

    \usepackage{amsmath,amsfonts,amsthm} % Math packages for equations

    \usepackage[svgnames, table]{xcolor} % Enabling colors by their 'svgnames'

    \usepackage[hang, small, labelfont=bf, up, textfont=it]{caption} % Custom captions under/above tables and figures

    \usepackage{booktabs} % Horizontal rules in tables

    \usepackage{lastpage} % Used to determine the number of pages in the document (for "Page X of Total")

    \usepackage{graphicx} % Required for adding images

    \usepackage{enumitem} % Required for customising lists
    \setlist{noitemsep} % Remove spacing between bullet/numbered list elements

    \usepackage{sectsty} % Enables custom section titles
    \allsectionsfont{\usefont{OT1}{phv}{b}{n}} % Change the font of all section commands (Helvetica)

    \usepackage{float} % Enabling colors by their 'svgnames'

    \usepackage{pdfpages}



    %----------------------------------------------------------------------------------------
    %	MARGINS AND SPACING
    %----------------------------------------------------------------------------------------

    \usepackage{geometry} % Required for adjusting page dimensions

    \geometry{
    	top=1cm, % Top margin
    	bottom=1.5cm, % Bottom margin
    	left=2cm, % Left margin
    	right=2cm, % Right margin
    	includehead, % Include space for a header
    	includefoot, % Include space for a footer
    	%showframe, % Uncomment to show how the type block is set on the page
    }

    \setlength{\columnsep}{7mm} % Column separation width

    %----------------------------------------------------------------------------------------
    %	FONTS
    %----------------------------------------------------------------------------------------

    \usepackage[T1]{fontenc} % Output font encoding for international characters
    \usepackage[utf8]{inputenc} % Required for inputting international characters

    \usepackage{XCharter} % Use the XCharter font

    %----------------------------------------------------------------------------------------
    %	HEADERS AND FOOTERS
    %----------------------------------------------------------------------------------------

    \usepackage{fancyhdr} % Needed to define custom headers/footers
    \pagestyle{fancy} % Enables the custom headers/footers

    \renewcommand{\headrulewidth}{0.0pt} % No header rule
    \renewcommand{\footrulewidth}{0.4pt} % Thin footer rule

    \renewcommand{\sectionmark}[1]{\markboth{#1}{}} % Removes the section number from the header when \leftmark is used

    %\nouppercase\leftmark % Add this to one of the lines below if you want a section title in the header/footer

    % Headers
    \lhead{} % Left header
    \chead{} % Center header - currently printing the article title
    \rhead{} % Right header

    % Footers
    \lfoot{} % Left footer
    \renewcommand{\headrulewidth}{0pt} % Suppress footer rule
    \cfoot{\footnotesize Page \thepage\ of \pageref{LastPage}} % footer, "Page 1 of 2"
    \renewcommand{\footrulewidth}{0pt} % Suppress footer rule
    \rfoot{} % footer, "Page 1 of 2"

    \fancypagestyle{firstpage}{ % Page style for the first page with the title
    	\fancyhf{}
    	\renewcommand{\headrulewidth}{0pt} % Suppress footer rule
        \cfoot{\footnotesize Page \thepage\ of \pageref{LastPage}} % footer, "Page 1 of 2"
        \renewcommand{\footrulewidth}{0pt} % Suppress footer rule
    }

    \fancypagestyle{zeropage}{ % Page style for the first page with the title
    	\fancyhf{}
    	\renewcommand{\headrulewidth}{0pt} % Suppress footer rule
        \cfoot{} % footer, "Page 1 of 2"
        \renewcommand{\footrulewidth}{0pt} % Suppress footer rule
    }

    %----------------------------------------------------------------------------------------
    %	TITLE SECTION
    %----------------------------------------------------------------------------------------
    %------------------------------------------------->> CHANGE HERE!
    \definecolor{MY_COLOR}{HTML}{'''

    B = colour

    C = r'''}

    \newcommand{\authorstyle}[1]{{\large\usefont{OT1}{phv}{b}{n}\color{black}#1}} % Authors style (Helvetica)

    \newcommand{\institution}[1]{{\footnotesize\usefont{OT1}{phv}{m}{sl}\color{Black}#1}} % Institutions style (Helvetica)

    \usepackage{titling} % Allows custom title configuration

    \newcommand{\HorRule}{\color{MY_COLOR}\rule{\linewidth}{1pt}} % Defines the gold horizontal rule around the title

    \pretitle{
        \vspace{-60pt} % Move the entire title section up
        \HorRule\vspace{10pt} % Horizontal rule before the title
        \fontsize{22}{26}\usefont{OT1}{phv}{b}{n}\selectfont % Helvetica
        \color{MY_COLOR} % Text colour for the title and author(s)
        \raggedright
    }

    %\usepackage{geometry}
    %\newgeometry{margin=1in} % Adjust the margin to 1 inch on all sides

    \posttitle{\par\vskip 19pt} % Whitespace under the title

    \preauthor{} % Anything that will appear before \author is printed

    \postauthor{ % Anything that will appear after \author is printed
    	\vspace{0pt} % Space before the rule
    	\par\HorRule % Horizontal rule after the title
    	\vspace{0pt} % Space after the title section
    }

    %----------------------------------------------------------------------------------------
    %	ABSTRACT
    %----------------------------------------------------------------------------------------

    \usepackage{lettrine} % Package to accentuate the first letter of the text (lettrine)
    \usepackage{fix-cm}	% Fixes the height of the lettrine

    \newcommand{\initial}[1]{ % Defines the command and style for the lettrine
    	\lettrine[lines=3,findent=4pt,nindent=0pt]{% Lettrine takes up 3 lines, the text to the right of it is indented 4pt and further indenting of lines 2+ is stopped
    		\color{MY_COLOR}% Lettrine colour
    		{#1}% The letter
    	}{}%
    }

    \usepackage{xstring} % Required for string manipulation

    \newcommand{\lettrineabstract}[1]{
    	\StrLeft{#1}{1}[\firstletter] % Capture the first letter of the abstract for the lettrine
    	\initial{\firstletter}\textbf{\StrGobbleLeft{#1}{1}} % Print the abstract with the first letter as a lettrine and the rest in bold
    }
    '''
    OUTPUT = A + B + C

    return OUTPUT

latex_head = r'''
\documentclass[10pt, a4paper, onecolumn]{article}
\input{structure.tex}
\begin{document}

'''

#---------------------------------------------------------------------->> MAIN - CODE STARTS HERE
print('Please enter the file with the answers in the form of a .csv file.')
input('Press >Enter< to continue...')

# GUI
root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
FILE_CONTENT = []
LATEX_OUTPUT = r''
LATEX_OUTPUT += latex_head
iCnt = -1
print('\n-->> Reading the file...')
ERROR_LOG = ''
try:
    #-->> is the file .csv?
    if file_path.endswith('.csv'):
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                FILE_CONTENT.append(row)
                iCnt = iCnt + 1
                time.sleep(0.1)
        print('\nNumber of answers read:',iCnt)
        time.sleep(1)

        print('\n-->> Starting the .csv to Latex conversion...')
        ans_num = len(FILE_CONTENT)
        
        #-->> main loop - create page after page of the latex document
        for iIi in range(1,ans_num):
            NUMBERE_OF_UNSUPORTED_CHAR = 0
            TMP_ANSWER = copy.deepcopy(FILE_CONTENT[iIi])
            TMP_ANSWER = list(filter(lambda x: x != '', TMP_ANSWER))
            starting_indx = 1
            if '@' in TMP_ANSWER[starting_indx]:
                starting_indx = 2

            pr_length = 80
            if len(TMP_ANSWER[starting_indx])>=pr_length:
                tmp_str_p = TMP_ANSWER[starting_indx][:pr_length] + '...'
            else:
                tmp_str_p = TMP_ANSWER[starting_indx]
            print('PAGE [',iIi,'/',ans_num-1,'] -',tmp_str_p)

            #-->> THE TITLE
            LATEX_OUTPUT += r'\title{' + escape_latex(TMP_ANSWER[starting_indx]) + r'} % The article title' + '\n'

            latex_abstract = TMP_ANSWER[starting_indx + 1]

            STATE = 0
            '''
            0 - no acknowledgements, complex answer
            1 - no acknowledgements, easy answer
            2 - acknowledgements, easy answer
            3 - acknowledgements, complex answer
            '''
            
            #-->> determine the type of answer - type of format
            try:
                num_author = int(TMP_ANSWER[starting_indx + 2])
                try:
                    num_af = int(TMP_ANSWER[starting_indx + 3])
                    STATE = 1
                    latex_thanks = ''
                except:
                    STATE = 0
            except:
                if TMP_ANSWER[starting_indx + 2] == 'More than 10':
                    STATE = 0
                    latex_thanks = ''
                else:
                    latex_thanks = TMP_ANSWER[starting_indx + 2]
                    try:
                        num_author = int(TMP_ANSWER[starting_indx + 3])
                        try:
                            num_af = int(TMP_ANSWER[starting_indx + 4])
                            STATE = 2
                        except:
                            STATE = 3
                    except:
                        STATE = 3
                        
            #-->> page format based on response type A -> contained more than 10 authors or 6 affiliations
            if STATE == 0 or STATE == 3:
                ERROR_LOG += '>> PAGE: [' + str(iIi) + '] contained more than 10 authors or 6 affiliations.\n   This program can fail to interpret the data correctly based on wrong user input. Check this page very carefully!\n\n'
                more_topic_name = TMP_ANSWER[starting_indx]
                more_abstract = TMP_ANSWER[starting_indx + 1]
                more_thanks = TMP_ANSWER[starting_indx + 2]
                more_is_student = TMP_ANSWER[-1]
                more_presenting_author = TMP_ANSWER[-2]
                more_authors = TMP_ANSWER[-3].split('\n')
                more_aff = TMP_ANSWER[-4].split('\n')

                #head
                LATEX_OUTPUT += r'\author{' +'\n' + r'\authorstyle{\noindent '

                for i in range(len(more_authors)):
                
                    #-->> RAW text preparation
                    tmp_afff_more = []
                    # Find all numbers enclosed by parentheses
                    numbers = re.findall(r'\(\d+\)', more_authors[i])

                    # Remove parentheses and enclosed numbers from the long string
                    for number in numbers:
                        tmp_afff_more.append(number)
                        more_authors[i] = more_authors[i].replace(number, '')

                    new_lst = []
                    for element in tmp_afff_more:
                        new_element = re.sub('[^0-9]', '', element)
                        new_lst.append(new_element)
                    tmp_afff_more = new_lst

                    #-->> Author
                    if more_presenting_author in more_authors[i]:
                        LATEX_OUTPUT += r'\underline{'+escape_latex(more_authors[i][0:-1])+'}'
                    else:
                        LATEX_OUTPUT += escape_latex(more_authors[i][0:-1])

                    LATEX_OUTPUT += r'\textsuperscript{'
                    for qq in range(len(tmp_afff_more)):
                        if qq != (len(tmp_afff_more) - 1):
                            LATEX_OUTPUT += escape_latex(tmp_afff_more[qq]) + ','
                        else:
                            LATEX_OUTPUT += escape_latex(tmp_afff_more[qq])
                    if i != (len(more_authors) - 1):
                        LATEX_OUTPUT += r'}, '
                    else:
                        LATEX_OUTPUT += r'}'

                LATEX_OUTPUT += r'}' + '\n'
                LATEX_OUTPUT += r'	\newline\newline' +'\n'


                #-->> Affiliations
                for i in range(len(more_aff)):
                    start = more_aff[i].find('(')  # find index of first '('
                    end = more_aff[i].find(')')  # find index of first ')'
                    if i == (len(more_aff) - 1):
                        if start != -1 and end != -1:  # if both '(' and ')' are found
                            LATEX_OUTPUT += r'	\textsuperscript{' + escape_latex(more_aff[i][start+1:end]) + r'}\institution{' + escape_latex(more_aff[i][end+2:]) + r'}' + '\n'
                        else:
                            LATEX_OUTPUT += r'	\institution{' + escape_latex(more_aff[i]) + r'}' + '\n'
                    else:
                        if start != -1 and end != -1:  # if both '(' and ')' are found
                            LATEX_OUTPUT += r'	\textsuperscript{' + escape_latex(more_aff[i][start+1:end]) + r'}\institution{' + escape_latex(more_aff[i][end+2:]) + r'}\\' + '\n'
                        else:
                            LATEX_OUTPUT += r'	\institution{' + escape_latex(more_aff[i]) + r'}\\' + '\n'
                LATEX_OUTPUT += r'}' + '\n\n'

                #-->> student award?
                if more_is_student == 'Yes':
                    LATEX_OUTPUT += r'\date{\fbox{Student Lecture}}' + '\n'
                else:
                    LATEX_OUTPUT += r'\date{}' + '\n'

                LATEX_OUTPUT += r'\maketitle' + '\n'
                LATEX_OUTPUT += r'\thispagestyle{firstpage} ' + '\n'

                #-->> ABSTRACT
                LATEX_OUTPUT += r'\lettrineabstract{' + escape_latex(more_abstract).replace('\n', r'\newline' + '\n' + r'\indent ') + r'}' + '\n'

                #-->> thanks
                if len(more_thanks)!=0:
                    LATEX_OUTPUT += r'\section*{\newline\newline\newline \normalsize Acknowledgements:}' + '\n' + escape_latex(more_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'
                else:
                    LATEX_OUTPUT += r'\section*{}' + '\n' + escape_latex(more_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'

                time.sleep(0.2)

            #-->> page format based on response type B -> contained less than 10 authors or 6 affiliations
            else:
                if STATE == 1:
                    author_indx = starting_indx + 4
                elif STATE == 2:
                    author_indx = starting_indx + 5

                #-->> SPECIAL CASE - one author, one affiliation
                if num_af==1 and num_author==1:
                    #head
                    LATEX_OUTPUT += r'\author{' +'\n' + r'\authorstyle{\noindent '
                    #-->> Authors
                    LATEX_OUTPUT += r'\underline{'+escape_latex(TMP_ANSWER[-3])+'}}\n'
                    LATEX_OUTPUT += r'	\newline\newline' +'\n'

                    LATEX_OUTPUT += r'	\institution{' + escape_latex(TMP_ANSWER[-2]) + r'}' + '\n'

                    LATEX_OUTPUT += r'}' + '\n\n'

                    #-->> student award?
                    if TMP_ANSWER[-1] == 'Yes':
                        LATEX_OUTPUT += r'\date{\fbox{Student Lecture}}' + '\n'
                    else:
                        LATEX_OUTPUT += r'\date{}' + '\n'

                    LATEX_OUTPUT += r'\maketitle' + '\n'
                    LATEX_OUTPUT += r'\thispagestyle{firstpage} ' + '\n'

                    #-->> ABSTRACT
                    LATEX_OUTPUT += r'\lettrineabstract{' + escape_latex(latex_abstract).replace('\n', r'\newline' + '\n' + r'\indent ') + r'}' + '\n'

                    #-->> thanks
                    if len(latex_thanks) != 0:
                        LATEX_OUTPUT += r'\section*{\newline\newline\newline \normalsize Acknowledgements:}' + '\n' + escape_latex(latex_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'
                    else:
                        LATEX_OUTPUT += r'\section*{}' + '\n' + escape_latex(latex_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'
                else:
                    #head
                    LATEX_OUTPUT += r'\author{' +'\n' + r'\authorstyle{\noindent '
                    #-->> Authors
                    if num_author==1:
                        presenting_author = 1
                    else:
                        presenting_author = TMP_ANSWER[-2]
                        tmp_string =''
                        for char in presenting_author:
                            if char.isnumeric():
                                tmp_string += char
                        presenting_author = int(tmp_string)

                    for i in range(author_indx,author_indx + num_author):
                        #Author
                        if (i-author_indx+1) == presenting_author:
                            LATEX_OUTPUT += r'\underline{'+escape_latex(TMP_ANSWER[i]).rstrip() +'}'
                        else:
                            LATEX_OUTPUT += escape_latex(TMP_ANSWER[i]).rstrip()

                        #affiliation
                        if num_af > 1 and num_author > 1:
                            LATEX_OUTPUT += r'\textsuperscript{'
                            tmp_aff = TMP_ANSWER[i + num_author + num_af].split(';')
                            for q in range(len(tmp_aff)):
                                tmp_string = ''
                                for char in tmp_aff[q]:
                                    if char.isnumeric():
                                        tmp_string += char
                                if q != (len(tmp_aff) -1):
                                    LATEX_OUTPUT += tmp_string + ','
                                else:
                                    LATEX_OUTPUT += tmp_string
                            if i != (author_indx + num_author -1):
                                LATEX_OUTPUT += r'}, '
                            else:
                                LATEX_OUTPUT += r'}'
                        else:
                            if i != (author_indx + num_author - 1):
                                LATEX_OUTPUT += r', '
                    LATEX_OUTPUT += r'}' + '\n'

                    LATEX_OUTPUT += r'	\newline\newline' +'\n'

                    #-->> affiliations
                    for i in range(num_af):
                        if num_af > 1:
                            if num_author > 1:
                                if i == (num_af - 1):
                                    LATEX_OUTPUT += r'	\textsuperscript{' + str(i+1) + r'}\institution{' + escape_latex(TMP_ANSWER[author_indx + num_author + i]) + r'}' + '\n'
                                else:
                                    LATEX_OUTPUT += r'	\textsuperscript{' + str(i+1) + r'}\institution{' + escape_latex(TMP_ANSWER[author_indx + num_author + i]) + r'}\\' + '\n'
                            else:
                                if i == (num_af - 1):
                                    LATEX_OUTPUT += r'	\textsuperscript{' + r'}\institution{' + escape_latex(TMP_ANSWER[author_indx + num_author + i]) + r'}' + '\n'
                                else:
                                    LATEX_OUTPUT += r'	\textsuperscript{' + r'}\institution{' + escape_latex(TMP_ANSWER[author_indx + num_author + i]) + r'}\\' + '\n'
                        else:
                            LATEX_OUTPUT += r'	\institution{' + escape_latex(TMP_ANSWER[author_indx + num_author + i]) + r'}' + '\n'
                    LATEX_OUTPUT += r'}' + '\n\n'

                    #-->> student awward?
                    if TMP_ANSWER[-1] == 'Yes':
                        LATEX_OUTPUT += r'\date{\fbox{Student Lecture}}' + '\n'
                    else:
                        LATEX_OUTPUT += r'\date{}' + '\n'

                    LATEX_OUTPUT += r'\maketitle' + '\n'
                    LATEX_OUTPUT += r'\thispagestyle{firstpage} ' + '\n'

                    #-->> ABSTRACT
                    LATEX_OUTPUT += r'\lettrineabstract{' + escape_latex(latex_abstract).replace('\n', r'\newline' + '\n' + r'\indent ') + r'}' + '\n'

                    #-->> thanks
                    if len(latex_thanks) != 0:
                        LATEX_OUTPUT += r'\section*{\newline\newline\newline \normalsize Acknowledgements:}' + '\n' + escape_latex(latex_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'
                    else:
                        LATEX_OUTPUT += r'\section*{}' + '\n' + escape_latex(latex_thanks) + '\n' + r'\newpage' + '\n\n\n\n\n\n\n\n\n\n'

            #-->> problems REPORT
            if NUMBERE_OF_UNSUPORTED_CHAR > 0:
                ERROR_LOG += '>> PAGE: [' + str(iIi) + '] contained ' + str(NUMBERE_OF_UNSUPORTED_CHAR) + " unsupported characters. Replacing by #!\n\n"
            NUMBERE_OF_UNSUPORTED_CHAR = 0
            if NUMBERE_OF_UNSUPORTED_GREEK > 0:
                ERROR_LOG += '>> PAGE: [' + str(iIi) + '] contained ' + str(NUMBERE_OF_UNSUPORTED_GREEK) + " unsupported Greek UNICODE characters. \n   No problem, replacing them with a latex alternatives.\n\n"
            NUMBERE_OF_UNSUPORTED_GREEK = 0
            time.sleep(0.2)
        LATEX_OUTPUT += r'\end{document}'

        time.sleep(0.2)
        print('\nDone!')

        time.sleep(1)
        print('\nProblems report:')
        time.sleep(1)
        if ERROR_LOG == '':
            print('No problems have occurred during the conversion :-)!')
        else:
            print(ERROR_LOG)

        time.sleep(2)
        print('\n-->> Creating the main.tex file to:')
        print(os.getcwd())
        # write the latex file
        with open("main.tex", "w", encoding="utf-8") as f:
            # Write the long string to the file object
            f.write(LATEX_OUTPUT)

        # Close the file object
        f.close()

        time.sleep(0.5)
        print('\n-->> Creating the format structure.tex file to:')
        print(os.getcwd())
        # Open a file object with write mode and the .tex file extension
        with open("structure.tex", "w", encoding="utf-8") as f:
            # Write the long string to the file object
            f.write(GenerateStructureFile('fdb813'))

        # Close the file object
        f.close()

        time.sleep(0.5)
        print('\nDone!')
        input('Press >Enter< to end this program...')

    else:
        print('Invalid file format. Please select a .csv file.')
        input('Press >Enter< to end this program...')

except:
    print('An error occurred during the conversion! Was the input a .csv file directly from google forms?')
    print('If so, delete the user input from the input csv file that is causing the problem.')
    print('Or contact kurka@ufe.cz')
    input('Press >Enter< to exit the program.')
