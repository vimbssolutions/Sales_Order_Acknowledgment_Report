from django.shortcuts import render

# Create your views here.
from django.shortcuts import render,get_object_or_404,redirect
from django.utils import timezone
from SOA_app.models import *
from django.http import HttpResponseRedirect, HttpResponse, FileResponse
from django.db.models import Q
from django.db import connection
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from SOA_app.SOA_Report import MyPrint
def SOA_Report(request):
    inc = 0
    chak = None
    a =list()
    next="False"
    ch=None
    # org = Common_Organization.objects.all()
    if request.method=="POST" and'pdf' in request.POST:
        # Create the HttpResponse object with the appropriate PDF headers.
        so_num=request.POST.get('Sonum')
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="My Users.pdf"'
        buffer = io.BytesIO()

        report = MyPrint(buffer, 'Letter')
        pdf = report.print_users(so_num)
        response.write(pdf)
        return response
    return render(request,'test.html')
