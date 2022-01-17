from django.shortcuts import render

# Create your views here.
from django.shortcuts import render
from . import report_download
from django.contrib import messages


def index(request):
    return render(request,'excel_report.html')

def download_data(request):
    if request.method == "POST":
        token_Get = str(request.POST.get("token"))
        date_Get = str(request.POST.get("date_val"))
        room_id_Get = str(request.POST.get("room_id"))
        download_file = report_download.download(room_id_Get,date_Get,token_Get,)
        #return render(request, 'excel_report.html')
        return render(request, 'excel_report.html',{'alert_flag': "True",'file_name':download_file})
    else:
        return render(request, 'excel_report.html')
