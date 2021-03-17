import re, openpyxl, keyboard
from sys import argv, exit
from os import path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QPushButton, QLabel

database = 'TextTable.xlsx'
file_path = input('Enter a msyt file. ')
file_content = open(file_path, 'r', encoding='utf-8').read()

if path.exists(database):
    text_xlsx = openpyxl.load_workbook('TextTable.xlsx')
    text_table = text_xlsx.get_sheet_by_name("Main")

print('\nPress F3 to add <c>')

app = QApplication(argv)

#النافذة الرئيسية
textbox_font = QtGui.QFont()
textbox_font.setPointSize(12)
label_font = QtGui.QFont()
label_font.setPointSize(14)

MainWindow = QMainWindow()
MainWindow.setFixedSize(326, 326)
MainWindow.setWindowTitle("Asgore msyt tool 1.32v")

msyt_text = QTextEdit(MainWindow)
msyt_text.setGeometry(QtCore.QRect(13, 13, 301, 123))
msyt_text.setFont(textbox_font)
sheet_text = QTextEdit(MainWindow)
sheet_text.setGeometry(QtCore.QRect(13, 193, 301, 123))
sheet_text.setFont(textbox_font)

next_button = QPushButton(MainWindow)
next_button.setGeometry(QtCore.QRect(210, 145, 93, 40))
next_button.setText("حفظ التغيير\nوالتالي")
save_button = QPushButton(MainWindow)
save_button.setGeometry(QtCore.QRect(105, 145, 93, 40))
save_button.setText("حفظ الملف")

per = QLabel(MainWindow)
per.setGeometry(QtCore.QRect(13, 145, 90, 40))
per.setFont(label_font)

def MsytToTxt(file_content):
    new_file_text, new_file_commands, new_file_dump = '', '', ''
    new_file_text_line, first_commands, last_commands = '', '', ''
    command_num = 0
    for line in file_content.split('\n'):
        if '- text:' in line:
            new_file_text_line += line.replace('      - text: ', '')
        elif '- control:' in line:
            new_file_text_line += '＜c' + str(command_num) + '＞'
            new_file_commands = new_file_commands.replace(']]', ']')
            new_file_commands += '[' + str(command_num) + ']]\n'
            command_num += 1
        elif '          ' in line:
            if 'animation' in line or 'sound' in line or 'sound2' in line or 'raw' in line:
                first_commands += '＜c' + str(command_num-1) + '＞'
                new_file_text_line = new_file_text_line.replace('＜c' + str(command_num-1) + '＞', '')
            elif 'auto_advance' in line or 'pause' in line or 'choice' in line or 'single_choice' in line:
                last_commands += '＜c' + str(command_num-1) + '＞'
                new_file_text_line = new_file_text_line.replace('＜c' + str(command_num-1) + '＞', '')
            
            new_file_commands = new_file_commands.replace(']]', ', ' + line.replace('          ', '') + ']]')
        else:
            if new_file_text_line:
                new_file_dump += '\t\t[-----------]\n'
                new_file_text += first_commands + new_file_text_line + last_commands + '\n'
                new_file_text_line, first_commands, last_commands = '', '', ''
            new_file_dump += line + '\n'
    
    new_file_content = '{\n'+new_file_text+'}\n\n' + '{\n'+new_file_commands+'}\n\n' + '{\n'+new_file_dump+'}'
    print(new_file_commands)
    return new_file_content

def TxtToMsyt(file_content):
    msyt_content_list = re.findall("\{\uffff(.*?)\uffff\}", file_content.replace('\n', '\uffff'))#for regex
    for i in range(len(msyt_content_list)): msyt_content_list[i] = msyt_content_list[i].replace('\uffff', '\n')
    
    TxtToMsyt.new_file_content = msyt_content_list[2]
    
    t = '\n' + msyt_content_list[0]
    text_list = t.split('\n')
    del text_list[0]
    
    def edit_line(line):
        if line[0] != '＜': line = '      - text: ' + line
        line = line.replace('\n', '\n      - text: ').replace('＞', '＞      - text: ')
        line = line.replace('＞      - text: ＜', '＞＜').replace('＞      - text: \n', '＞\n')
        line = line.replace('＞      - text: ', '＞\n      - text: ').replace('＜', '\n＜')
        line = line.replace('""', '')
        TxtToMsyt.new_file_content = TxtToMsyt.new_file_content.replace('\t\t[-----------]', line, 1)
    
    list(map(edit_line, text_list))
    
    commands_list = re.findall("\[(.*?)\]", msyt_content_list[1])
    for i in range(len(commands_list)):
        j = '\n      - control:' + commands_list[i].replace(str(i)+', ', ', ').replace(', ', '\n          ')
        TxtToMsyt.new_file_content = TxtToMsyt.new_file_content.replace('＜c' + str(i) + '＞', j)
    
    TxtToMsyt.new_file_content = TxtToMsyt.new_file_content.replace('\n      - text: \n', '\n')
    return TxtToMsyt.new_file_content.replace('\n\n\n      - control:', '\n      - control:').replace('\n\n      - control:\n', '\n      - control:\n')

def script():
    global Text_content
    if sheet_text.toPlainText():
        Text_content = Text_content.replace(text_list[script.current_item-1], sheet_text.toPlainText())
        text_list[script.current_item-1] = sheet_text.toPlainText()
    
    msyt_text.setPlainText(text_list[script.current_item])
    sheet_text.setPlainText('')
    
    if path.exists(database):
        for item in re.split('＜(.*?)＞', text_list[script.current_item]):
            if item:
                for cell in range(2, len(text_table['A'])+1):
                    if text_table['A'+str(cell)].value == item:
                        sheet_text.setPlainText(sheet_text.toPlainText()+text_table['B'+str(cell)].value)
            
    if script.current_item == len(text_list)-1:
        per.setText(f"{centences_num} \ {centences_num}")
        script.current_item = 0
    else:
        script.current_item += 1
        per.setText(f"{centences_num} \ {script.current_item}")

script.current_item = 0

def typeC():
    c = 0
    if '＞' in msyt_text.toPlainText():
        while '＜c'+str(c)+'＞' not in msyt_text.toPlainText(): c += 1
        for i in range(msyt_text.toPlainText().count('＞')):
            if '＜c'+str(c+i)+'＞' not in sheet_text.toPlainText() and '＜c'+str(c+i)+'＞' in msyt_text.toPlainText():
                keyboard.write('＜c'+str(c+i)+'＞')
                break

if file_path.endswith('.msyt'):
    Text_content = MsytToTxt(file_content)
    msyt_content_list = re.findall("\{\uffff(.*?)\uffff\}", Text_content.replace('\n', '\uffff'))#for regex
    for i in range(len(msyt_content_list)): msyt_content_list[i] = msyt_content_list[i].replace('\uffff', '\n')
    centences_num = msyt_content_list[0].count('\n')+1
    
    t = '\n' + msyt_content_list[0]
    text_list = t.split('\n')
    del text_list[0]
    
    script()

next_button.clicked.connect(lambda: script())
save_button.clicked.connect(lambda: open(file_path, 'w', encoding='utf-8').write(TxtToMsyt(Text_content)))
keyboard.on_press_key("F3", lambda _: typeC())

MainWindow.show()
exit(app.exec_())