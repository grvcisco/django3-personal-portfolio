
import xlwt
from django.http import HttpResponse
from django.shortcuts import render, get_object_or_404
from .models import Blog
from django.core.mail import EmailMessage
from io import BytesIO
from django.conf import settings

def all_blogs(request):
    blogs = Blog.objects.order_by('-date')
    return render(request, 'blog/all_blogs.html', {'blogs':blogs})

def detail(request, blog_id):
    blog = get_object_or_404(Blog, pk=blog_id)
    return render(request, 'blog/detail.html', {"id":blog})


def export_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="blogs.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Blogs')

     # Sheet header, first row
    row_num = 0
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    columns = ['Title ', 'Date Created', 'Description', ]
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()

    rows = Blog.objects.all().values_list('title', 'date',  'desc')

    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    wb.save(response)


    # Send excel as attachment
    excelfile = BytesIO()
    wb.save(excelfile)

   """ email = EmailMessage()
    email.subject = 'Test subject'
    email.body = 'Testing email attachments in django'
    email.from_email = settings.EMAIL_HOST_USER
    email.to = ['gsrivas3@cisco.com']
    email.attach('blogs.xls', excelfile.getvalue(), 'application/ms-excel')
    email.send()
    """

    return response
