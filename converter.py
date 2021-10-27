from tkinter.ttk import Progressbar
from tkinter import filedialog, scrolledtext, messagebox, ttk
from tkinter import *
import PyPDF2
import os
import re
import contextlib
import win32com.client
import img2pdf
import fitz
import PIL
import subprocess
from pdfrw import PdfReader


HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<title>{title}</title>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
</head>
<body>
{divs}
</body>
</html> 
"""

file = ''
fold = ''

# описание О программе
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

# отработка кнопки Отчитска выбра
def clicked_cln():
    global file, fold
    txt_file.delete(1.0, END)
    txt_fold.delete(1.0, END)
    file = ''
    fold = ''
    bar['value'] = 0
    lbl_3.configure(text="")
    lbl_3_2.configure(text="")
    lbl_3_1.configure(text="")

# Экранирование "/" в пути файла
def change2(file_in):
    file_out = re.sub(r'/', r'\\\\', file_in)
    return file_out

# Экранирование "\" в пути файла
def change(file_in):
    file_out = re.sub(r'\\', r'/', file_in)
    return file_out

# Редактируем именя файлов, удаляем старое расширение
def cut_name(file_in):
    file_out = file_in[:file_in.rfind('.')]
    return file_out

# Обрезка пути файла
def cut_dir(file_in):
    file_out = file_in[:file_in.rfind('/')]
    return file_out

#Размер PDF Документа
def size_pdf(file):
    pdf = PdfReader(file)
    return pdf.pages[0].MediaBox[2]

#Добавление водяного знака в PDF файл
def add_vatermark_pdf(file):
    size = size_pdf(file)
    barcode_file = r"\\172.30.0.2\ящики\install\PDF&html converter\logo.jpg"
    if int(float(size)) <= 600:
        fitz.image_rectangle = fitz.Rect(495, 0, 595, 35)
    else:
        if int(float(size)) >= 840:
            fitz.image_rectangle = fitz.Rect(740, 0, 840, 35)
    file_handle = fitz.open(file)
    first_page = file_handle[0]
    first_page.insertImage(fitz.image_rectangle, filename=barcode_file)
    file_handle.save(file, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

# pdf to html convert
def pdf_html(in_file):
    doc = fitz.open(in_file)
    title = doc.metadata["title"]
    divs = ""
    for i in range(doc.pageCount):
        div = doc.getPageText(i, "htm")
        divs += div
    html = HTML_TEMPLATE.format(title=title, divs=divs)
    with open(change2(cut_name(in_file)) + '.htm', 'w') as f:
        f.write(html)
        f.close()

# Преобразование doc to pdf
def doc2pdf(in_file):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.options.ConfirmConversions = 0
        word.visible = 0
        wb = word.Documents.Open(change2(in_file))
        wb.SaveAs2(change2(cut_name(in_file))+'.pdf', FileFormat=17)
        wb.Close()
        word.Quit()
    except:
        return 1

# Переименование файла из временного 123 в необходимое имя и перемещение в его каталог
def repl_pdf(self):
    name = self[len(cut_dir(self)) + 1:len(cut_name(self))]
    os.rename(r'C:/temp/123.pdf', r'C:/temp/' + name + '.pdf')
    os.replace(r'C:/temp/' + name + '.pdf', change2(cut_name(self)) + '.pdf')

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
def excel2pdf(file_in):
    try:
        pages = []
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(change2(file_in))
        if len(wb.Worksheets) > 1:
            for i in range(len(wb.Worksheets)):
                ws = wb.Worksheets[i]
                ws.ExportAsFixedFormat(0, 'C:\\temp\\123_' + str(i) + '.pdf', False, False)
                pages.append('C:\\temp\\123_' + str(i) + '.pdf')
            if len(pages) > 1:
                pdf_add_page(pages)
                repl_pdf(file_in)
                for page in pages:
                    os.remove(page)
        else:
            wb.ActiveSheet.SaveAs('C:\\temp\\123', FileFormat=57)
            repl_pdf(file_in)
        wb.Close()
        excel.Quit()
    except:
        wb.ActiveSheet.SaveAs('C:\\temp\\123', FileFormat=57)
        repl_pdf(file_in)
        wb.Close()
        excel.Quit()
        os.remove('C:\\temp\\123_0.pdf')
    finally:
        return 1

# tiff to pdf конвертация(с учетом книжной или альбомной ориентации листа)
def tif2pdf(self):
    img = PIL.Image.open(self)
    if img.size[0] < 2600:
        a4_page_size = [img2pdf.in_to_pt(8.3), img2pdf.in_to_pt(11.7)]
        layout_function = img2pdf.get_layout_fun(a4_page_size)
        pdf = img2pdf.convert(self, layout_fun=layout_function)
    else:
        a4_rotate_size = [img2pdf.in_to_pt(11.7), img2pdf.in_to_pt(8.3)]
        layout_function = img2pdf.get_layout_fun(a4_rotate_size)
        pdf = img2pdf.convert(self, layout_fun=layout_function)
    with open(cut_name(self) + '.pdf', 'wb') as f:
        f.write(pdf)

# Удаление исходного файла если стоит чек на удаление
def state_dell_file(self):
    if chk_state_dell.get() == 1:
        os.remove(self)

# Конвертация файла в pdf основная
def convert_file_pdf(in_file):
    # if chk_state_pdf.get() == 1:
    if in_file.endswith('.docx') or in_file.endswith('.doc') or in_file.endswith('.rtf'):
        doc2pdf(in_file)
        state_dell_file(in_file)
    if in_file.endswith('.xlsx') or in_file.endswith('.xls'):
        excel2pdf(in_file)
        state_dell_file(in_file)
    if in_file.endswith('.tif'):
        tif2pdf(in_file)
        state_dell_file(in_file)
    if chk_state_vaterm.get() == 1:
        add_vatermark_pdf(change2(cut_name(in_file))+'.pdf')

# конвертация файла в html
def con_file_html(in_file):
    if in_file.endswith('.pdf'):
        pdf_html(in_file)
        state_dell_file(in_file)
    else:
        convert_file_pdf(in_file)
        pdf_html(cut_name(in_file) + '.pdf')
        os.remove(cut_name(in_file) + '.pdf')
        state_dell_file(in_file)

# отработка конвертации 1 файла
def con_file(in_file):
    if chk_state_pdf.get() == 1:
        convert_file_pdf(in_file)
    if chk_state_html.get() == 1:
        con_file_html(in_file)

# счетчик файлов в каталоге, проценты для progressbar
def file_count(path):
    cmd = 'DIR \"%s\" /A-D /B /S | FIND /C /V ""' % path
    try:
        from subprocess import DEVNULL
    except ImportError:
        DEVNULL = os.open(os.devnull, os.O_RDWR)
    count = int(subprocess.check_output(cmd, shell=True, stdin=DEVNULL, stderr=DEVNULL))
    return count

# конвертирование папки в pdf
def con_folder(self):
    folder = []
    countfiles = file_count(self)
    proc = 100 / countfiles
    count = 0
    for i in os.walk(self):
        folder.append(i)
    for address, dirs, files in folder:
        for fil in files:
            count = count + 1
            lbl_3_1.configure(text=("Обработано: "+ str(count) +" из " + str(countfiles)))
            lbl_3.configure(text="")
            lbl_3_2.configure(text="")
            bar['value'] += proc
            if len(fil) <= 62:
                lbl_3.configure(text=fil)
            else:
                lbl_3.configure(text=fil[:62])
                lbl_3_2.configure(text=fil[62:123])
            window.update()
            con_file(change(address + '/' + fil))
    return messagebox.showinfo('Внимание!!', 'Конвертация успешно завершена!')

# Отработка нажатия кнопки Конвертирования
def clicked_con():
    if chk_err() == 1:
        if len(file) > 0:
            con_file(file)
        if len(fold) > 0:
            if chk_state_pdf.get() == 1 or chk_state_html.get() == 1:
                con_folder(fold)

# Создание интерфейса
window = Tk()
window.title("Pdf & Htm converter")
window.geometry('465x500')
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
file_btn.grid(row=2, column=0, sticky=E + W)
txt_file = scrolledtext.ScrolledText(window, width=55, height=1)
txt_file.grid(column=0, row=1, sticky=E)
lbl_2 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_2.grid(column=0, row=3)

# Выбор папки для конвертирования(кнопка)
fold_btn = Button(window, text="Выберите папку", command=clicked_fold)
fold_btn.grid(row=5, column=0, sticky=E + W)
txt_fold = scrolledtext.ScrolledText(window, width=55, height=1)
txt_fold.grid(column=0, row=4, sticky=E)
lbl_3_1 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_3_1.grid(column=0, row=6)

#Проггресс бар
bar = Progressbar(window, length=300, style='black.Horizontal.TProgressbar', maximum=100, value=0)
bar.grid(column=0, row=7)
lbl_3 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_3.grid(column=0, row=9)
lbl_3_2 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_3_2.grid(column=0, row=10)
cln_btn = Button(window, text="Очистить!", command=clicked_cln)
cln_btn.grid(row=11, column=0)

# Чекбоксы для выбора режимов работы
chk_state_pdf = IntVar()
chk_state_pdf.set(0)
chk_state_html = IntVar()
chk_state_html.set(0)
chk_state_vaterm = IntVar()
chk_state_vaterm.set(0)
chk_pdf = Checkbutton(window, text='PDF', var=chk_state_pdf, bg='#808080')
chk_pdf.grid(column=0, row=12, sticky=W)
chk_pdf = Checkbutton(window, text='HTML', var=chk_state_html, bg='#808080')
chk_pdf.grid(column=0, row=13, sticky=W)

# Конверировать(кнопка)
ttk.Style().configure("TButton", padding=10, relief="RAISED", background="#ccc")
con_btn = ttk.Button(window, text="Конвертировать!", command=clicked_con)
con_btn.grid(row=14, column=0)
lbl_5 = Label(window, text="", font=("Arial Bold", 10), bg='#808080')
lbl_5.grid(column=0, row=15)

#Добавление водяного знака в PDF
chk_vaterm = Checkbutton(window, text='Добавить водяной знак', var=chk_state_vaterm, bg='#808080')
chk_vaterm.grid(column=0, row=16, sticky=W)

# Удаление файлов(чекбокс)
lbl2 = Label(window, text="Внимание! Выбор удалит исходные файлы!",
             font=("Arial Bold", 10), bg='#808080', fg='#dcdcdc')
lbl2.grid(column=0, row=17)
chk_state_dell = IntVar()
chk_state_dell.set(0)
chk_dell = Checkbutton(window, text='Удаление исходных файлов! ',
                       var=chk_state_dell, bg='#808080')
chk_dell.grid(column=0, row=18, sticky=W)
window.mainloop()