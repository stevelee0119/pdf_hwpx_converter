from reportlab.pdfgen import canvas

def create_sample_pdf(path):
    c = canvas.Canvas(path)
    text = "This is a sample PDF file with multiple words and spaces."
    c.drawString(100, 750, text)
    c.showPage()
    c.save()

if __name__ == "__main__":
    create_sample_pdf("sample.pdf")
