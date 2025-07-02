import os
import time
import logging
import traceback
from tkinter import filedialog, Tk
from progress.bar import IncrementalBar
import fitz  # PyMuPDF


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
            return

        log.info(f"Found {len(items_to_process)} folders to process")

        # Обработка папок
        with IncrementalBar('Processing', max=len(items_to_process)) as bar:
            for folder in items_to_process:
                try:
                    bar.message = f"Processing {folder[:20]}..."
                    process_folder(base_dir, folder, result_dir)
                    log.info(f"Successfully processed: {folder}")
                except Exception as e:
                    log.error(f"Error processing {folder}: {str(e)}")
                    log.error(traceback.format_exc())
                finally:
                    bar.next()

        log.info(f"Program completed in {time.time() - start_time:.2f} seconds")

    except Exception as e:
        log.error(f"Critical error: {str(e)}")
        log.error(traceback.format_exc())


def process_folder(base_dir, folder, result_dir):
    folder_path = os.path.join(base_dir, folder)
    files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.pdf'))
    ]

    if not files:
        raise ValueError("No valid image or PDF files found")

    doc = fitz.open()
    has_pages = False

    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                # Обработка изображений
                with fitz.open(file_path) as img_doc:
                    pdf_bytes = img_doc.convert_to_pdf()
                    with fitz.open("pdf", pdf_bytes) as img_pdf:
                        doc.insert_pdf(img_pdf)
                        has_pages = True
            elif file.lower().endswith('.pdf'):
                # Обработка PDF
                with fitz.open(file_path) as pdf_doc:
                    if pdf_doc.page_count > 0:
                        doc.insert_pdf(pdf_doc)
                        has_pages = True
        except Exception as e:
            logging.warning(f"Failed to process {file}: {str(e)}")
            continue

    if not has_pages:
        raise ValueError("No valid pages were added to the document")

    output_path = os.path.join(result_dir, f"{folder}.pdf")
    doc.save(output_path)
    doc.close()


if __name__ == "__main__":
    main()
