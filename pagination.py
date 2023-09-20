import PyPDF2
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
import io
import pdfscript

def add_text_to_pdf(input_file_path, output_file_path):
    # Открываем исходный PDF-файл для чтения
    with open(input_file_path, 'rb') as file:
        # Создаем объект PdfFileReader для работы с PDF-файлом
        pdf = PyPDF2.PdfFileReader(file)
        # Получаем количество страниц в PDF-файле
        num_pages = pdf.getNumPages()

        # Создаем буфер BytesIO для временного хранения новых страниц с текстом
        packet = io.BytesIO()
        # Создаем объект canvas для рисования на странице
        can = canvas.Canvas(packet, pagesize=landscape(A4))

        # Цикл для перебора каждой страницы PDF-файла
        for page_num in range(num_pages):
            if page_num % 2 == 0:
                pdfscript.n = pdfscript.n + 1
                # Получаем текущую страницу PDF
                page = pdf.getPage(page_num)

                # Получаем размеры страницы
                page_width = float(page.mediaBox[2])
                page_height = float(page.mediaBox[3])

                # Добавляем текст "hello" в нижний правый угол страницы
                text = str(pdfscript.n)
                text_width = can.stringWidth(text, "Helvetica", 10)
                x = (page_width - text_width) - 10  # Отступ от правого края страницы
                y = 10  # Отступ от нижнего края страницы
                can.drawString(x, y, text)
          
                
            # Завершаем рисование на текущей странице и переходим к следующей
            can.showPage()

        # Сохраняем рисование на всех страницах в буфере
        can.save()

        # Перемещаем указатель буфера в начало
        packet.seek(0)

        # Создаем новый объект PdfFileReader из буфера, который содержит страницы с текстом "hello"
        new_pdf = PyPDF2.PdfFileReader(packet)

        # Создаем объект PdfFileWriter для создания нового PDF-файла с добавленным текстом
        output = PyPDF2.PdfFileWriter()

        # Цикл для объединения исходных страниц с текстом "hello"
        for page_num in range(num_pages):
            # Получаем текущую страницу из исходного PDF
            page = pdf.getPage(page_num)
            # Объединяем текущую страницу с текстом "hello" из объекта new_pdf
            page.mergePage(new_pdf.getPage(page_num))
            # Добавляем объединенную страницу в объект output
            output.addPage(page)

        # Открываем новый файл для записи
        with open(output_file_path, "wb") as output_file:
            # Записываем объединенные страницы в новый PDF-файл с именем "output.pdf"
            output.write(output_file)
