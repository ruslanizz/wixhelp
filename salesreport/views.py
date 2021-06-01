from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render

from salesreport.models import SalesReport
from salesreport.services import handle_sales_report


def salesreport_page(request):
    mydata = {}

    if request.method == 'POST' and request.FILES.get('1c_file') and request.FILES.get('csv_file'):

        uploaded_1c_file = request.FILES['1c_file']
        print('Загружен файл ',uploaded_1c_file.name)
        print('Размер файла ', uploaded_1c_file.size)

        uploaded_csv_file = request.FILES['csv_file']
        print('Загружен файл ', uploaded_csv_file.name)
        print('Размер файла ', uploaded_csv_file.size)

        # response = HttpResponseRedirect(redirect_to='salesreport_log/', content_type='text/csv')
        # response = HttpResponse(content_type='text/csv')
        # response['Content-Disposition'] = 'attachment; filename="ready_to_wix.csv"'

        mydata, mylog = handle_sales_report(uploaded_1c_file, uploaded_csv_file)

        # SalesReport.objects.all().delete()
        # oneentry=SalesReport.objects.create(log=mylog)

        # mydata.to_csv(path_or_buf=response, index=False)
        mydata.to_csv(path_or_buf='salesreport/static/ready_to_wix.csv', index=False)



        # return response
        # mytext='aaaaaaaaaaaaa'

        return render(request, 'salesreport_log.html', context={'data': mylog})

    return render(request, 'index.html')


# def salesreport_log_page(request):
#     if request.method == 'POST':
#         print('AAAAAAAAAAAA !!!!!')
#     log=SalesReport.objects.first().log
#     return render(request, 'salesreport_log.html', context={'data':log})
