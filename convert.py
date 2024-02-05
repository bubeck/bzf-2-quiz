#!/usr/bin/python

import argparse
import subprocess
import sys
from openpyxl import Workbook

def convert_pdf(filename):
    p = subprocess.run(["pdftotext", "-raw", "-f", "7", filename, "-"],
                       capture_output=True,
                       check=True, encoding="utf8")
    return p.stdout

def strip_header(text):
    pages = text.split("\f")
    pages2 = []
    for page in pages:
        lines = page.split("\n")
        page = "\n".join(lines[2:])
        pages2.append(page)
    return "\n".join(pages2)

def find_question(lines, line_no, number):
    q = []
    if lines[line_no] != f'{number}':
        # There is text after the number
        q.append(lines[line_no][len(f'{number}'):])
    while True:
        line_no = line_no + 1
        if lines[line_no].startswith("A "):
            break
        q.append(lines[line_no])
    return (" ".join(q), line_no)

def find_answers(lines, line_no):
    a = []
    answer_no = "A"
    #if lines[line_no] != f'{answer_no}':
        # There is text after the number
        #a.append(lines[line_no][len(1:]))
    #    :
    while True:
        line_no = line_no + 1
        if lines[line_no].startswith("A "):
            break
        q.append(lines[line_no])
    return (" ".join(q), line_no)

def split_questions(text):
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
            #print(line_start)
            #print(line_no)
            #print("")
            line_start = line_no
            number = number + 1
        line_no = line_no + 1
        
    n = f'{number-1}'
    lines[line_start] = lines[line_start][len(n)+1:] 
    q.append("\n".join(lines[line_start:line_no]))
    
    return q

def remove_empty_lines(questions):
    for i in range(len(questions)):
        q = []
        for line in questions[i].split("\n"):
            if line.strip() != "":
                q.append(line.strip())
        questions[i] = "\n".join(q)
    return questions

def join_question(questions):
    for i in range(len(questions)):
        q1 = questions[i].split("\n")
        q2 = []
        for j in range(len(q1)):
            if q1[j].startswith("A ") or q1[j] == "A":
                q2.append(" ".join(q1[0:j]))
                q2 = q2 + q1[j:]
                #print(q2)
                questions[i] = "\n".join(q2)
                break

    return questions
    
def split_answers(questions):
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
                #print(q2)
        q1[answer_start] = q1[answer_start][2:]
        q2.append(" ".join(q1[answer_start:]).strip())
        
        questions[i] = "\n".join(q2)

    return questions
    
parser = argparse.ArgumentParser(
                    #prog='ProgramName',
                    description='Konvertiere BZF-2 Fragenkatalog PDF in Aiken',
                    #epilog='Text at the bottom of help')
    )
parser.add_argument('filename', help="Filename of BZF-2 PDF from Bundesnetzagentur")           # positional argument
args = parser.parse_args()

txt = convert_pdf(args.filename)
txt = strip_header(txt)
questions = split_questions(txt)
questions = remove_empty_lines(questions)
questions = join_question(questions)
questions = split_answers(questions)

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
    print(q2[0])
    row = [str(q_no), q2[0]]
    
    for i in range(1, len(q2)):
        answer = chr(ord("A")-1+i)
        print(f'{answer}. {q2[i]}')
        row.append(q2[i])
        if i == 1:
            row.append("1")
        else:
            row.append("")
    ws.append(row)
    
    print("ANSWER: A\n")
    
#print(txt)
wb.save("bzf-2-quiz.xslx")
