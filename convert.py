#!/usr/bin/python

"""
This script converts a PDF question catalog from the Bundesnetzagentur (BZF-2)
to the Aiken quiz format and saves the results in an Excel file (.xlsx).

The data structure 'questions' is a list of strings, where each string
contains a question and its corresponding answer options. The structure looks like this:

questions = [
    "Question 1\nAnswer 1A\nAnswer 1B\nAnswer 1C",
    "Question 2\nAnswer 2A\nAnswer 2B\nAnswer 2C",
    ...
]

Note: The answer options are arranged in the original PDF order.
"""

import argparse
import subprocess
from openpyxl import Workbook

def convert_pdf(filename):
    """
    Converts a PDF file to raw text using the 'pdftotext' external program.
    
    Parameters:
    - filename (str): The filename of the BZF-2 PDF from Bundesnetzagentur.

    Returns:
    - str: The raw text extracted from the PDF.
    """
    p = subprocess.run(["pdftotext", "-raw", "-f", "7", filename, "-"],
                       capture_output=True,
                       check=True, encoding="utf8")
    return p.stdout

def strip_header(text):
    """
    Removes the header part from the extracted text.
    
    Parameters:
    - text (str): The raw text extracted from the PDF.

    Returns:
    - str: The text with the header removed.
    """
    pages = text.split("\f")
    pages2 = []
    for page in pages:
        lines = page.split("\n")
        page = "\n".join(lines[2:])
        pages2.append(page)
    return "\n".join(pages2)

def split_questions(text):
    """
    Splits the text into individual questions based on question numbers.
    
    Parameters:
    - text (str): The text with the header removed.

    Returns:
    - list: A list of strings, each containing a question and its answer options.
    """
    lines = text.split("\n")
    number = 2
    line_no = 0
    q = []
    line_start = line_no
    for line_no in range(len(lines)):
        if lines[line_no].startswith(f'{number}'):
            n = f'{number-1}'
            lines[line_start] = lines[line_start][len(n)+1:] 
            q.append("\n".join(lines[line_start:line_no]))
            line_start = line_no
            number = number + 1
        line_no = line_no + 1
        
    n = f'{number-1}'
    lines[line_start] = lines[line_start][len(n)+1:] 
    q.append("\n".join(lines[line_start:line_no]))
    
    return q

def remove_empty_lines(questions):
    """
    Removes empty lines from each question.
    
    Parameters:
    - questions (list): A list of strings, each containing a question and its answer options.

    Returns:
    - list: A list of strings with empty lines removed from each question.
    """
    for i in range(len(questions)):
        q = []
        for line in questions[i].split("\n"):
            if line.strip() != "":
                q.append(line.strip())
        questions[i] = "\n".join(q)
    return questions

def join_question(questions):
    """
    Joins the question and all associated answer options into a single string.
    
    Parameters:
    - questions (list): A list of strings, each containing a question and its answer options.

    Returns:
    - list: A list of strings with questions and associated answer options joined.
    """
    for i in range(len(questions)):
        q1 = questions[i].split("\n")
        q2 = []
        for j in range(len(q1)):
            if q1[j].startswith("A ") or q1[j] == "A":
                q2.append(" ".join(q1[0:j]))
                q2 = q2 + q1[j:]
                questions[i] = "\n".join(q2)
                break

    return questions

def split_answers(questions):
    """
    Splits the answer options for each question and formats them.
    
    Parameters:
    - questions (list): A list of strings, each containing a question and its answer options.

    Returns:
    - list: A list of strings with formatted answer options for each question.
    """
    for i in range(len(questions)):
        q1 = questions[i].split("\n")
        q2 = [q1[0]]
        answer_no = "B"
        answer_start = 1
        for j in range(1,len(q1)):
            if q1[j].startswith(f'{answer_no} ') or q1[j] == answer_no:
                q1[answer_start] = q1[answer_start][2:]
                q2.append(" ".join(q1[answer_start:j]).strip())
                answer_start = j
                answer_no = chr(ord(answer_no)+1)
        q1[answer_start] = q1[answer_start][2:]
        q2.append(" ".join(q1[answer_start:]).strip())
        
        questions[i] = "\n".join(q2)

    return questions

# Argument processing via the command line with argparse
parser = argparse.ArgumentParser(description='Convert BZF-2 PDF question catalog to Aiken')
parser.add_argument('filename', help="Filename of BZF-2 PDF from Bundesnetzagentur")           
args = parser.parse_args()

# Convert PDF to text and remove unnecessary header lines
txt = convert_pdf(args.filename)
txt = strip_header(txt)

# Extract and format questions
questions = split_questions(txt)
questions = remove_empty_lines(questions)
questions = join_question(questions)
questions = split_answers(questions)

# Create Excel file and write data
wb = Workbook()
ws = wb.active
title = ["No", "Question"]
for i in range(1,11):
    title.append(f'Answer{i}')
    title.append(f'AnswerCorrect{i}')
title = title + ["Hint", "Explanation", "Category"]
ws.append(title)

q_no = 0
for q in questions:
    q_no = q_no + 1
    q2 = q.split("\n")
    row = [str(q_no), q2[0]]
    
    for i in range(1, len(q2)):
        answer = chr(ord("A")-1+i)
        row.append(q2[i])
        if i == 1:
            row.append("1")
        else:
            row.append("")
    ws.append(row)

# Save Excel file
wb.save("bzf-2-quiz.xslx")
