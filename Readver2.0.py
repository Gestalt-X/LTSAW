# -*- coding: utf-8 -*-
"""
Rewrite the code completely for better performance, being more user-fridendly and more readable.
This is my first formal project. I'll make it!
Learn more about me on my Github account: Gestalt-X. Welcome to visit my repository: LTSAW. 
@author: Gestalt-X
"""
# Part0 Find the path

def lock(): 
    global Path
    key = int(input(r'Please enter "1" or "2" : ')) 
    if key == 1:
        Path = input('Please enter the path of the document: ')
        return readtxt(Path)
    elif key == 2:
        Path = input('Please enter the path of the document: ')
        return readxlsx(Path)
    else:
        raise TypeError(r'Either "1" nor "2"')

# Part1  Read and Write .txt

def readtxt(path):
    with open(path,'r',encoding='utf8') as f:
        lines = f.readlines()
        print(lines)
        f.close()
        branch = input('Enter "Y" to write the document or "N" to quit the programme : ')
        if branch == 'Y':
            return writetxt(Path)
        elif branch == 'N':
            exit()
        else:
            raise TypeError(r'Either "Y" nor "N"')
   
def writetxt(path):
    with open(path,'a',encoding='utf8') as f:
         txt = input('Please enter the words you want to add : ')
         f.write(txt)
         f.close()
    return readtxt(Path)

# Part2  Read .xlsx
    
import xlrd

def readxlsx(path):
    xl = xlrd.open_workbook(path)
    sheet = xl.sheets()[0]
    data = []
    for i in range(0,sheet.ncols):
        data.append(list(sheet.col_values(i)))
    print(data)
    return data

# Main functions as greetings

import datetime

def main():
    
    # Words
    cont1 = "Greetings!"
    now = datetime.datetime.now()
    cont2 = "Nice to meet you here at "+now.strftime("%Y-%m-%d %H:%M:%S")
    cont3 = "Press \"1\" to read txt; Press \"2\" to read xlsx."
    
    # Size 
    size1 = len(cont1)
    size2 = len(cont2)
    size3 = len(cont3)
    len_head = max(size1+4,size2+4,size3+4)
    
    # Frame
    head1 = '+' + '-'*(len_head-2) + '+'
    head2 = '|' + ' '*(len_head-2) + '|'
    print(head1,head2,sep='\n')
    
    # Output
    output1 = '|' + cont1.center(len_head-2) + '|'
    output2 = '|' + cont2.center(len_head-2) + '|'
    output3 = '|' + cont3.center(len_head-2) + '|'
    print(output1,output2,output3,sep='\n')
    print(head2,head1,sep='\n')

    return lock()
              
# Start main at last

if __name__ == '__main__':
    main()
