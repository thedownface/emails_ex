from fileinput import filename
from django.shortcuts import HttpResponse,render
from scrapy.crawler import CrawlerProcess
from email_ex_scrapy.email_ex_scrapy.spiders.email_ex import EmailExtractor
from scrapy.utils.project import get_project_settings

def e_scrapy(request):
    name='emailex111'
    response = HttpResponse(content_type = 'applications/ms-excel')
    response['Content-Disposition'] = 'attachment; filename = "emails.xlsx"'
    if "GET" == request.method:
        return render(request, 'e_scrapy/scrapy.html', {})
    else:
        excel_file = request.FILES["excel_file"]
    
        process = CrawlerProcess(get_project_settings())
        EmailExtractor.read_excel(excel_file)

        process.crawl(EmailExtractor)
        process.start()
        process.stop()
        