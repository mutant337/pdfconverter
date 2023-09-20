from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter


packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=landscape(letter))
can.drawString(700, 10, "Hello world")
can.save()

#move to the beginning of the StringIO buffer
packet.seek(0)

# create a new PDF with Reportlab
new_pdf = PdfFileReader(packet)
# read your existing PDF
existing_pdf = PdfFileReader(open(r"C:\Users\kostiantyn.dzhelalov\Desktop\pdfconverter\samples docs\A4-booklet-landscape.en.pdf", "rb"))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.pages[0]
page.merge_page(new_pdf.pages[0])
output.add_page(page)
# finally, write "output" to a real file
output_stream = open(r"C:\Users\kostiantyn.dzhelalov\Desktop\pdfconverter\samples docs\res.pdf", "wb")
output.write(output_stream)
output_stream.close()