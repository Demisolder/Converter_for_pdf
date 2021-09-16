import os
from logging import FileHandler
import PIL.Image
from PyPDF2 import PdfFileMerger, PdfFileReader
import shutil
from tkinter import filedialog
from tkinter import *
import time
from progress.bar import IncrementalBar
import logging
import traceback
import sys


logfile = 'log_1.log'

log = logging.getLogger("my_log")
log.setLevel(logging.INFO)
FH = logging.FileHandler(logfile, encoding='utf-8')
basic_formater = logging.Formatter('%(asctime)s : [%(levelname)s] : %(message)s')
FH.setFormatter(basic_formater)
log.addHandler(FH)

def img2pdf(fname): # конвертация jpg в pdf, копирование в папку in_pdf
    filename = fname
    name = filename.split('.')[0]  # имя файла до точки перед расширением
    im = PIL.Image.open(filename)
    #os.chdir(dirName1)
    newfilename = ''.join([name,'.pdf'])# путь, включая имя и расширение
    PIL.Image.Image.save(im , newfilename , "PDF" , resolution=100.0)
    #print("processed successfully: {}".format(newfilename))
    os.remove(fname)

log.info("start program")
start_time = time.time()
try:
    root = Tk()
    root.withdraw()   #запрашиваем директорию начала скрипта
    dirName1 = filedialog.askdirectory()
    #dirName1 = 'C:/Users/1/Desktop/Новая папка (5)'

    os.chdir(dirName1)
    arr = os.listdir()
    #print(arr)
    bar = IncrementalBar('Countdown', max = len(arr))
    if not os.path.exists(dirName1 + '/result_pdf/'):
        # проверяем есть ли папка
        os.makedirs(dirName1 + '/result_pdf/')  # если нет - создаём
    #создаём список имён без расширения в папке result_pdf
    res_list = (os.listdir(os.path.join(dirName1 + '/result_pdf/')))
    res_list2 = list(res_list)
    for i in range(len(res_list)):
        name = res_list[i]
        res_list2[i] = name.split('.')[0]
        # имя файла до точки перед расширением
    #print(res_list2)
    nabor =set(res_list2)# c набором поиск элемента быстрее
    for i in range(len(arr)):
        bar.next()
        #time.sleep(1)
        if arr[i] == "result_pdf":#проверка на папку для результатов
            continue
        # проверяем есть ли элемент arr[i] в папке результатов.
        elif arr[i] in nabor:
            print('папка ' + arr[i] + ' уже обработана'+"--- %s seconds ---" % (time.time() - start_time))
            log.info("папка  " + arr[i]+ " уже обработана"+"--- %s seconds ---" % (time.time() - start_time))
            continue
        # основной код конвертации
        else:
            for l in range(5): #даём 5 попыток на переподключение
                try:
                    dirName2 = os.path.join(dirName1, arr[i])
                    os.chdir(dirName2)# перешли в директорию подпапки
                    dirName3 = os.path.join(os.environ['USERPROFILE'] + '\Desktop' + '\in_pdf')
                    # проходим по файлам подпапки arr[i]
                    # конвертация отдельных файлов в PDF
                    for fname in os.listdir(dirName2):
                        if not os.path.exists(dirName3 + '/' + arr[i]): #создаём папку in_pdf, если её нет
                            os.makedirs(dirName3 + '/' + arr[i])#создаём папку in_pdf_arr[i]
                        # если jpg - конвертируем в папку in_pdf_arr[i]
                        if fname.endswith(".jpg"):
                            shutil.copyfile(dirName1 + '/' + arr[i] + '/' + fname , dirName3 + '/' + arr[i] + '/' + fname)
                            os.chdir(dirName3 + '/' + arr[i])
                            img2pdf(fname)
                        # pdf документы мы копируем в папку in_pdf_arr[i]
                        elif fname.endswith(".pdf"):
                            t=1
                            newfilename=fname
                            while os.path.exists(dirName3 + '/' + arr[i] + '/' + newfilename):
                                #пока существует такое имя, мы прибавляем к нему сумматор
                                newfilename= str(t) + fname
                                t=t+1
                            shutil.copyfile(dirName1 + '/' + arr[i] + '/' + fname , dirName3 + '/' + arr[i] + '/' + newfilename)
                    subfolder = (os.listdir(dirName3 + '/' + arr[i]))
                    # проименовали массив с именами файлов в подпапке
                    os.chdir(dirName3 + '/' + arr[i])
                    # соединение отдельных файлов в 1 PDF
                    merger = PdfFileMerger()
                    EOF_MARKER = b'%%EOF' #обработка исключений EOF
                    for pdf in subfolder:
                        with open(pdf, 'rb') as f:
                            contents = f.read()
                        # check if EOF is somewhere else in the file
                        if EOF_MARKER in contents:
                            # we can remove the early %%EOF and put it at the end of the file
                            contents = contents.replace(EOF_MARKER , b'')
                            contents = contents + EOF_MARKER
                        else:
                            # Some files really don't have an EOF marker
                            # In this case it helped to manually review the end of the file
                            print(contents[-8:])  # see last characters at the end of the file
                            # printed b'\n%%EO%E'
                            contents = contents[:-6] + EOF_MARKER

                        with open(pdf.replace('.pdf' , '') + '_fixed.pdf' , 'wb') as f:
                            if b'\r%%EOF' in contents:
                                contents = contents.replace(b'\r%%EOF', b'\%%EOF')
                                if b'\r\%%EOF' in contents:
                                    contents = contents.replace(b'\r\%%EOF' , b'\n%%EOF')
                            f.write(contents)

                            newPDF = f.name

                        #print (pdf)
                        merger.append(PdfFileReader(newPDF, 'rb'))  # (text_pdf_file, strict=False))
                        #merger.close()
                    os.chdir(dirName1 + '/result_pdf/')
                    merger.write(arr[i] + ".pdf")
                    merger.close()
                    shutil.rmtree(dirName3 + '/' + arr[i])
                    print("создан файл " + arr[i] + ".pdf"+"--- %s seconds ---" % (time.time() - start_time))
                    log.info("создан файл  " + arr[i] + ".pdf"+"--- %s seconds ---" % (time.time() - start_time))
                except:
                    #print("конечная ошибка - " + traceback.format_exc())
                    print ("ошибка конвертации файла " + str(arr[i]) + " попытка № "+ str(l) + traceback.format_exc())

                    log.info("ошибка конвертации файла " + arr[i])
                    if l == 4 :
                        os.makedirs(dirName1 +'/result_pdf/' + arr[i])  # создаём папку arr[i] в папке результатов
                        for fname in os.listdir(dirName2):
                            shutil.copyfile(dirName1 + '/' + arr[i] + '/' + fname ,dirName1 + '/result_pdf/' + arr[i] + '/' + fname)
                    continue
                else:
                    break


    bar.finish()
    #shutil.rmtree(dirName3)
except :
    frame = traceback.extract_tb(sys.exc_info()[2])
    line_no = str(frame[0]).split()[4]
    ## вызываем функцию записи ошибки и передаем в нее номер строки с ошибкой
    log.error("конечная ошибка - " + traceback.format_exc())
    print("конечная ошибка - " + traceback.format_exc())
log.info("end program" + dirName1)
print("--- %s seconds ---" % (time.time() - start_time))
log.info("--- %s seconds ---" % (time.time() - start_time))





