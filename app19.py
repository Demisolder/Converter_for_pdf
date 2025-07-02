import os
import time
import logging
import traceback
from tkinter import filedialog, Tk, Toplevel, Label, ttk, Button
from tkinter import messagebox
from progress.bar import IncrementalBar
import fitz  # PyMuPDF
import threading
from PIL import Image
import io


# Настройка логирования
def setup_logging():
    logfile = 'conversion.log'
    logger = logging.getLogger("pdf_converter")
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler(logfile, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s : [%(levelname)s] : %(message)s')
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    return logger


# Функция сжатия изображений
def compress_image(file_path, max_size=(1000, 1000), quality=85):
    try:
        img = Image.open(file_path)
        img.thumbnail(max_size, Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality)
        return buffer.getvalue()
    except Exception as e:
        logging.warning(f"Failed to compress {file_path}: {str(e)}")
        return None


def process_folder(base_dir, folder, result_dir, cancel_event):
    folder_path = os.path.join(base_dir, folder)
    files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.pdf'))
    ]

    if not files:
        raise ValueError("No valid image or PDF files found")

    doc = fitz.open()
    has_pages = False
    skipped_files = []

    for file in files:
        if cancel_event.is_set():
            doc.close()
            raise ValueError("Processing cancelled by user")

        file_path = os.path.join(folder_path, file)
        try:
            # Проверка на поврежденный файл
            if not os.path.getsize(file_path) > 0:
                logging.warning(f"File {file} is empty or corrupted")
                skipped_files.append(file)
                continue

            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                # Сжатие изображения
                compressed_data = compress_image(file_path)
                if compressed_data:
                    with fitz.open("jpg", compressed_data) as img_doc:
                        pdf_bytes = img_doc.convert_to_pdf()
                        with fitz.open("pdf", pdf_bytes) as img_pdf:
                            doc.insert_pdf(img_pdf)
                            has_pages = True
                else:
                    skipped_files.append(file)
                    continue
            elif file.lower().endswith('.pdf'):
                # Обработка PDF
                with fitz.open(file_path) as pdf_doc:
                    if pdf_doc.page_count > 0:
                        doc.insert_pdf(pdf_doc)
                        has_pages = True
                    else:
                        logging.warning(f"PDF {file} has no pages")
                        skipped_files.append(file)
        except Exception as e:
            logging.warning(f"Failed to process {file}: {str(e)}")
            skipped_files.append(file)
            continue

    if not has_pages:
        doc.close()
        raise ValueError("No valid pages were added to the document")

    output_path = os.path.join(result_dir, f"{folder}.pdf")
    doc.save(output_path)
    doc.close()
    return skipped_files


def process_folder_in_thread(base_dir, folder, result_dir, progress_label, progress_bar, index, total, cancel_event, skipped_files_list):
    try:
        skipped_files = process_folder(base_dir, folder, result_dir, cancel_event)
        logging.info(f"Successfully processed: {folder} (Progress: {index + 1}/{total})")
        skipped_files_list.extend([(folder, f) for f in skipped_files])
    except Exception as e:
        logging.error(f"Error processing {folder}: {str(e)}")
        logging.error(traceback.format_exc())
    finally:
        progress_bar['value'] = index + 1
        progress_label.config(text=f"Обработка: {folder[:20]}... ({index + 1}/{total})")
        progress_bar.update()


def main():
    log = setup_logging()
    log.info("Program started")
    start_time = time.time()

    try:
        # Выбор директории
        root = Tk()
        root.withdraw()
        base_dir = filedialog.askdirectory()
        if not base_dir:
            log.info("No directory selected. Exiting.")
            messagebox.showinfo("Завершено", "Директория не выбрана.")
            root.destroy()
            return

        os.chdir(base_dir)
        result_dir = os.path.join(base_dir, 'result_pdf')
        os.makedirs(result_dir, exist_ok=True)

        # Получаем список уже обработанных файлов
        processed_files = {
            os.path.splitext(f)[0]
            for f in os.listdir(result_dir)
            if f.endswith('.pdf')
        }

        # Фильтруем только директории для обработки
        items_to_process = [
            item for item in os.listdir(base_dir)
            if os.path.isdir(item)
            and item != 'result_pdf'
            and item not in processed_files
        ]

        if not items_to_process:
            log.info("No new folders to process")
            messagebox.showinfo("Завершено", "Нет новых папок для обработки.")
            root.destroy()
            return

        log.info(f"Found {len(items_to_process)} folders to process")

        # Создаем окно прогресса
        progress_window = Toplevel()
        progress_window.title("Обработка файлов")
        progress_window.geometry("400x200")
        progress_window.resizable(False, False)

        Label(progress_window, text="Обработка папок...").pack(pady=10)
        progress_label = Label(progress_window, text="")
        progress_label.pack(pady=5)
        progress_bar = ttk.Progressbar(progress_window, maximum=len(items_to_process), mode='determinate')
        progress_bar.pack(pady=10, padx=20, fill='x')

        # Кнопка отмены
        cancel_event = threading.Event()
        cancel_button = Button(progress_window, text="Отмена", command=cancel_event.set)
        cancel_button.pack(pady=10)

        # Список для хранения пропущенных файлов
        skipped_files_list = []

        # Обработка папок в отдельных потоках
        threads = []
        with IncrementalBar('Processing', max=len(items_to_process)) as bar:
            for i, folder in enumerate(items_to_process):
                thread = threading.Thread(
                    target=process_folder_in_thread,
                    args=(base_dir, folder, result_dir, progress_label, progress_bar, i, len(items_to_process), cancel_event, skipped_files_list)
                )
                threads.append(thread)
                thread.start()
                bar.next()

        # Ожидание завершения всех потоков
        for thread in threads:
            thread.join()

        # Проверка на отмену
        if cancel_event.is_set():
            log.info("Processing was cancelled by user")
            messagebox.showwarning("Отмена", "Обработка была отменена пользователем.")
            progress_window.destroy()
            root.destroy()
            return

        # Формирование сообщения о завершении
        completion_message = f"Обработка завершена за {time.time() - start_time:.2f} секунд."
        if skipped_files_list:
            completion_message += "\n\nПропущенные файлы:\n"
            for folder, file in skipped_files_list:
                completion_message += f"{folder}/{file}\n"

        log.info(f"Program completed in {time.time() - start_time:.2f} seconds")
        progress_window.destroy()
        messagebox.showinfo("Завершено", completion_message)
        root.destroy()

    except Exception as e:
        log.error(f"Critical error: {str(e)}")
        log.error(traceback.format_exc())
        messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
        root.destroy()


if __name__ == "__main__":
    main()
