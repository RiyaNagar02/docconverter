import os
from django.shortcuts import render
from django.http import FileResponse
from django.conf import settings
from .utils import convert_pdf_to_docx


def upload_pdf(request):
    if request.method == "POST":
        pdf_file = request.FILES["pdf"]

        pdf_path = os.path.join(settings.MEDIA_ROOT, "input.pdf")
        docx_path = os.path.join(settings.MEDIA_ROOT, "output.docx")

        with open(pdf_path, "wb+") as f:
            for chunk in pdf_file.chunks():
                f.write(chunk)

        convert_pdf_to_docx(pdf_path, docx_path)

        return FileResponse(
            open(docx_path, "rb"),
            as_attachment=True,
            filename="converted.docx"
        )

    return render(request, "upload.html")
from django.shortcuts import render

# Create your views here.
