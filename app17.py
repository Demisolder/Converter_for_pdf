import os
from logging import FileHandler
import shutil
from tkinter import filedialog
from tkinter import *
import time
from progress.bar import IncrementalBar
import logging
import traceback
import sys
import fitz

logfile = 'log_1.log'
Num = 1
log = logging.getLogger("my_log")
log.setLevel(logging.INFO)
FH = logging.FileHandler(logfile, encoding='utf-8')
basic_formater = logging.Formatter('%(asctime)s : [%(levelname)s] : %(message)s')
FH.setFormatter(basic_formater)
log.addHandler(FH)

log.info("start program")
start_time = time.time()
try:
    root = Tk()
    root.withdraw()  # запрашиваем директорию начала скрипта
    dirName1 = filedialog.askdirectory()
    os.chdir(dirName1)
    arr = os.listdir()
    bar = IncrementalBar('Countdown', max=len(arr))

    if not os.path.exists(dirName1 + '/result_pdf/'):
        os.makedirs(dirName1 + '/result_pdf/')  # если нет - создаём

    res_list = (os.listdir(dirName1 + '/result_pdf/'))
    for i in range(len(res_list)):
        name = res_list[i]
        res_list[i] = name.split('.')[0]  # имя файла до точки перед расширением
    nabor = set(res_list)  # c набором поиск элемента быстрее

    for i in range(len(arr)):
        bar.next()
        if arr[i] == "result_pdf":  # проверка на папку для результатов
            continue
        elif arr[i] in nabor:
            print('папка ' + arr[i] + ' уже обработана' + "--- %s seconds ---" % (time.time() - start_time))
            continue
        else:
            try:
                dirName2 = os.path.join(dirName1, arr[i])
                os.chdir(dirName2)
                spisok = os.listdir(dirName2)
                doc = fitz.open()  # new PDF
                has_pages = False  # флаг, есть ли валидные страницы

                for img in spisok:
                    try:
                        if img.lower().endswith((".jpg", ".jpeg", ".png", ".bmp")):
                            imgdoc = fitz.open(img)  # open image as a document
                            pdfbytes = imgdoc.convert_to_pdf()  # make a 1-page PDF of it
                            imgpdf = fitz.open("pdf", pdfbytes)
                            doc.insert_pdf(imgpdf)  # insert the image PDF
                            imgdoc.close()
                            imgpdf.close()
                            has_pages = True
                        elif img.lower().endswith(".pdf"):
                            try:
                                pdfbytes = fitz.open(img)  # open PDF
                                if pdfbytes.page_count > 0:  # проверяем, есть ли страницы
                                    doc.insert_pdf(pdfbytes)  # insert the PDF
                                    has_pages = True
                                pdfbytes.close()
                            except:
                                print("ошибка конвертации файла " + str(arr[i]) + " " + img)
                                log.info("ошибка конвертации файла " + str(arr[i]) + " " + img)
                                continue
                    except:
                        print("ошибка обработки файла " + str(arr[i]) + " " + img)
                        log.info("ошибка обработки файла " + str(arr[i]) + " " + img)
                        continue

                os.chdir(dirName1 + '/result_pdf/')
                if has_pages and doc.page_count > 0:  # проверяем, есть ли страницы перед сохранением
                    output_file_path = arr[i] + ".pdf"
                    doc.save(output_file_path)
                    print("создан файл №%d " % Num + " " + arr[i] + ".pdf" + "--- %s seconds ---" % (
                                time.time() - start_time))
                    Num += 1
                else:
                    print("папка " + arr[i] + " не содержит валидных изображений или PDF-файлов")
                    log.info("папка " + arr[i] + " не содержит валидных изображений или PDF-файлов")

                doc.close()
                continue
            except:
                frame = traceback.extract_tb(sys.exc_info()[2])
                line_no = str(frame[0]).split()[4]
                print("ошибка на стадии сохранения слияния файлов - " + traceback.format_exc())
                log.error("ошибка на стадии сохранения слияния файлов - " + traceback.format_exc())
                continue
    bar.finish()
except:
    frame = traceback.extract_tb(sys.exc_info()[2])
    line_no = str(frame[0]).split()[4]
    log.error("конечная ошибка - " + traceback.format_exc())
    print("конечная ошибка - " + traceback.format_exc())

log.info("end program" + dirName1)
print("--- %s seconds ---" % (time.time() - start_time))
log.info("--- %s seconds ---" % (time.time() - start_time))
