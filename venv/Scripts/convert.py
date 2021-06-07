from docx2pdf import convert
# from tkinter.ttk import Checkbutton
from tkinter import filedialog, scrolledtext, messagebox, ttk
from tkinter import *
import PyPDF2
import os
import re
import contextlib
import win32com.client
import PIL
import sys
import reportlab
from reportlab.pdfgen import canvas

file = ''
fold = ''

def about():
    messagebox.showinfo('О программе:', 'Данный конвертер преобразует '
                        'файлы Word, Excel и Tif в PDF или Html.\n\n Автор: ScaLseR')

# Отработка нажатия кнопки открытия файла
def clicked_file():
    global file
    clicked_cln()
    file = filedialog.askopenfilename()
    txt_file.insert(INSERT, file)

# Отработка нажатия кнопки открытия папки
def clicked_fold():
    global fold
    clicked_cln()
    fold = filedialog.askdirectory()
    txt_fold.insert(INSERT, fold)

# Проверка на ошибки формы
def chk_err():
    if file == '' and fold == '':
        messagebox.showwarning('Внимание!!', 'Не выбран файл или папка для конвертирования!')
        return 0
    if chk_state_pdf.get() == 0 and chk_state_html.get() == 0:
        messagebox.showwarning('Внимание!!', 'Не выбран режим конвертирования!')
        return 0
    if chk_state_pdf.get() == 1 and chk_state_html.get() == 1:
        messagebox.showwarning('Внимание!!', 'Выбраны оба режима конвертирования! '
                               'Выберите только 1 из режимов!')
        return 0
    return 1

# Отчитска выбра
def clicked_cln():
    global file, fold
    txt_file.delete(1.0, END)
    txt_fold.delete(1.0, END)
    file = ''
    fold = ''

# Экранирование "\" в пути файла
def resub(file_in):
    file_out = re.sub(r'/', r'\\\\', file_in)
    return file_out

# Редактируем именя файлов, удаляем старое расширение
def cut_name(file_in):
    file_out = file_in[:file_in.rfind('.')]
    # print('file_out',file_out)
    return file_out

# Обрезка пути файла
def cut_dir(file_in):
    file_out = file_in[:file_in.rfind('/')]
    # print('file_out',file_out)
    return file_out

# Преобразование doc to docx
def doc2docx(in_file):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0
        wb = word.Documents.Open(resub(in_file))
        wb.SaveAs2(in_file + 'x', FileFormat=16)
        wb.Close()
        word.Quit()
    except:
        return 0
    else:
        return 1

# конвертация doc,docx в html
def doc2html(in_file):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0
        wb = word.Documents.Open(resub(in_file))
        wb.SaveAs2(resub(cut_name(in_file))+'.html', FileFormat=8)
        wb.Close()
        word.Quit()
    except:
        return 0
    else:
        return 1

# конвертация xls,xlsx в html
def xls2html(in_file):
    try:
        excel = win32com.client.Dispatch("Word.Application")
        excel.visible = 0
        wb = excel.Workbooks.Open('d:\\123.xlsx')
        wb.ExportAsFixedFormat(0, 'd:\\123.html', False, False)
        wb.Close()
        excel.Quit()
    except:
        return 0
    else:
        return 1

# Переименование файла из временного 123 в необходимое имя и перемещение в его каталог
def repl_pdf(in_file):
    try:
        name = in_file[len(cut_dir(in_file)) + 1:len(cut_name(in_file))]
        os.rename(r'C:/temp/123.pdf', r'C:/temp/' + name + '.pdf')
        os.replace(r'C:/temp/' + name + '.pdf', resub(cut_name(in_file)) + '.pdf')
    except:
        return 0

# Сборка в 1 pdf из нескольких
def pdf_add_page(pdf_files_list):

    with contextlib.ExitStack() as stack:
        pdf_merger = PyPDF2.PdfFileMerger()
        files = [stack.enter_context(open(pdf, 'rb')) for pdf in pdf_files_list]
        for f in files:
            pdf_merger.append(f)
        with open(r'C:/temp/123.pdf', 'wb') as f:
            pdf_merger.write(f)

# конвертация Xls в pdf
def excel2pdf(in_file):
    try:
        pages = []
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = 0
        wb = excel.Workbooks.Open(resub(in_file))
        if len(wb.Worksheets) > 1:
            for i in range(len(wb.Worksheets)):
                ws = wb.Worksheets[i]
                # ws.Visible = 1
                ws.ExportAsFixedFormat(0, 'C:\\temp\\123_' + str(i) + '.pdf', False, False)
                pages.append('C:\\temp\\123_' + str(i) + '.pdf')
            pdf_add_page(pages)
            repl_pdf(in_file)
            for page in pages:
                os.remove(page)
        else:
            ws = wb.Worksheets[0]
            wb.SaveAs('C:\\temp\\123', FileFormat=57)
            repl_pdf(in_file)
        wb.Close()
        excel.Quit()
    except:
        print('Ошибка конвертации ЕXCEL')
    finally:
        return 1

def tif2pdf(in_file):
    # outPDF = canvas.Canvas(cut_name(in_file)+'.pdf', pageCompression=1)
    # img = PIL.Image.open(in_file)
    # for page in range(img.n_frames):
    #    img.seek(page)
     #   imgPage = reportlab.lib.utils.ImageReader(img)
     #   outPDF.drawImage(imgPage, 0, 0, 595, 841)
     #   if page < img.n_frames:
     #       outPDF.showPage()
    # outPDF.save()
    # img.close()


# Удаление исходного файла если стоит чек на удаление
def state_dell_file(in_file):
    if chk_state_dell.get() == 1:
        os.remove(in_file)

# Конвертация файла основная
def convert_file(in_file):
    if chk_state_pdf.get() == 1:
        if in_file.endswith('.docx'):
            convert(in_file)
            state_dell_file(in_file)
        if in_file.endswith('.doc'):
            if doc2docx(in_file) == 1:
                convert(in_file+'x')
                os.remove(in_file+'x')
                state_dell_file(in_file)
        if in_file.endswith('.xlsx') or in_file.endswith('.xls'):
            excel2pdf(in_file)
            state_dell_file(in_file)
        if in_file.endswith('.tif'):
            tif2pdf(in_file)
            state_dell_file(in_file)
    if chk_state_html.get() == 1:
        if in_file.endswith('.docx') or in_file.endswith('.doc'):
            doc2html(in_file)
        if in_file.endswith('.xls') or in_file.endswith('.xlsx'):
            xls2html(in_file)
        state_dell_file(in_file)

# Отработка нажатия кнопки Конвертирования
def clicked_con():
    if chk_err() == 0:
        print('Продолжаем выполнение чека на ошибки!!!!')
    else:
        if len(file) > 0:
            convert_file(file)
        if len(fold) > 0:
            print(fold)


# Создание интерфейса
window = Tk()
window.title("Pdf & Html converter")
window.geometry('465x415')
window.configure(bg='#808080')

# Меню с описанием "О программе"
menu = Menu(window)
menu.add_command(label='Справка', command=about)
window.config(menu=menu)

# Описание
lbl1 = Label(window, text="Выберите файл или папку для конвертирования:",
             font=("Arial Bold", 10), bg='#808080', fg='#dcdcdc')
lbl1.grid(column=0, row=0)
lbl_1 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_1.grid(column=0, row=1)

# Выбор файла для конвертирования(кнопка)
file_btn = Button(window, text="Выберите файл", command=clicked_file)
file_btn.grid(row=2, column=0, sticky=E+W)
txt_file = scrolledtext.ScrolledText(window, width=55, height=1)
txt_file.grid(column=0, row=1, sticky=E)
lbl_2 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_2.grid(column=0, row=3)

# Выбор папки для конвертирования(кнопка)
fold_btn = Button(window, text="Выберите папку", command=clicked_fold)
fold_btn.grid(row=5, column=0, sticky=E+W)
txt_fold = scrolledtext.ScrolledText(window, width=55, height=1)
txt_fold.grid(column=0, row=4, sticky=E)

lbl_3 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_3.grid(column=0, row=7)
cln_btn = Button(window, text="Очистить!", command=clicked_cln)
cln_btn.grid(row=8, column=0)

# Чекбоксы для выбора режимов работы
chk_state_pdf = IntVar()
chk_state_pdf.set(0)
chk_state_html = IntVar()
chk_state_html.set(0)
chk_pdf = Checkbutton(window, text='PDF', var=chk_state_pdf, bg='#808080')
chk_pdf.grid(column=0, row=9, sticky=W)
chk_pdf = Checkbutton(window, text='HTML', var=chk_state_html, bg='#808080')
chk_pdf.grid(column=0, row=10, sticky=W)

# Конверировать(кнопка)
ttk.Style().configure("TButton", padding=10, relief="RAISED", background="#ccc")
con_btn = ttk.Button(window, text="Конвертировать!", command=clicked_con)
con_btn.grid(row=11, column=0)

lbl_5 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_5.grid(column=0, row=12)

# Удаление файлов(чекбокс)
lbl2 = Label(window, text="Внимание! Выбор удалит исходные файлы!",
             font=("Arial Bold", 10), bg='#808080', fg='#dcdcdc')
lbl2.grid(column=0, row=13)
chk_state_dell = IntVar()
chk_state_dell.set(0)
chk_dell = Checkbutton(window, text='Удаление исходных файлов! ',
                       var=chk_state_dell, bg='#808080')
chk_dell.grid(column=0, row=14, sticky=W)

window.mainloop()