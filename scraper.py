import requests
from bs4 import BeautifulSoup
import os, shutil
from urllib.parse import urljoin
from pypdf import PdfReader, PdfWriter
import xlsxwriter

URL = 'https://www.wfis.uni.lodz.pl/strefa-studenta/plany-zajec/' #hardcoded schedules URL
headers ={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:136.0) Gecko/20100101 Firefox/136.0'}
folder = './schedules'
export = folder + '/export'


print('Creating "schedules/" folder...')
if not os.path.exists(folder):
    os.mkdir(folder)
else:
    for file in os.listdir(folder):  # Clear schedules folder
        file_path = os.path.join(folder, file)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to remove file %s, error: %s' % (file, e))

if not os.path.exists(export):
    os.mkdir(export)

print('Searching for pdfs...')
page = requests.get(URL, headers=headers) #getting page

soup = BeautifulSoup(page.content, 'html.parser') #parsing page content

my_div = soup.find(id="slidedown-80588-162589") #finding div with my schedules

pdfs = []
for a in my_div('a', href=True):
    link = a['href']
    if link[-4:] == '.pdf':
        print(f'Found pdf file: {link}. Saving to the schedules folder...')
        filename = os.path.join(folder, link.split('/')[-1])
        with open(filename, 'wb') as file:
            file.write(requests.get(urljoin(URL, link,)).content) #finding and saving the .pdf files
            pdfs.append(file.name)


def format_excel(wb, b3, c3, d3, e3, b4, c4, d4, e4):
    worksheet = wb.add_worksheet()
    worksheet.set_column('A:E', 40)
    worksheet.set_row(1, 50)
    worksheet.set_row(2, 150)
    worksheet.set_row(3, 150)
    worksheet.write("B2", "8:15 - 10:30", header)
    worksheet.write("C2", "10:45 - 13:00", header)
    worksheet.write("D2", "13:45 - 16:00", header)
    worksheet.write("E2", "16:15 - 18:30", header)
    worksheet.write("A3", "Saturday", header)
    worksheet.write("A4", "Sunday", header)
    worksheet.write("B3", b3, body)
    worksheet.write("C3", c3, body)
    worksheet.write("D3", d3, body)
    worksheet.write("E3", e3, body)
    worksheet.write("B4", b4, body)
    worksheet.write("C4", c4, body)
    worksheet.write("D4", d4, body)
    worksheet.write("E4", e4, body)

workbook = xlsxwriter.Workbook(export + '/Schedule.xlsx')
header = workbook.add_format({'bold': True, 'font_size': 30})
header.set_align('center')
header.set_align('vcenter')
header.set_border(6)
body = workbook.add_format({'font_size': 15})
body.set_align('center')
body.set_align('vcenter')
body.set_border(1)
body.set_text_wrap(True)
writer = PdfWriter() #Creating pdf
for pdf in pdfs:
    reader = PdfReader(pdf)
    if pdf.find('Terminy') > 0:
        writer.add_page(reader.pages[0])
        print('Creating a new "Schedules.pdf" file. Adding first page...')
    else:
        for page in reader.pages:
            if page.extract_text(0).find('INFORMATYKA II st.') > 0 and page.extract_text(0).find('I rok 2 sem.') > 0:
                print("Adding saturday and sunday to the Schedule.pdf")
                writer.add_page(page)
                page1 = reader.pages[page.page_number + 1]
                writer.add_page(page1)
                pages = [page, page1]
                my_classes = []
                c_b3, c_c3, c_d3, c_e3, c_b4, c_c4, c_d4, c_e4 = '', '', '', '', '', '', '', ''
                substr = page.extract_text().split(' \n \n')
                for sub in substr:
                    if not sub.find('GR 1') and not sub.find('SZTUCZNA')and not sub.isspace() and sub.find('godz'):
                        for _ in range(2):
                            sub = sub.removeprefix('\n')
                            sub = sub.removeprefix(' \n')
                            sub = sub.removeprefix(' ')
                        if sub.find('8.15') > 0: c_b3 = sub
                        elif sub.find('10.45') > 0: c_c3 = sub
                        elif sub.find('13.45') > 0: c_d3 = sub
                        elif sub.find('16.15') > 0: c_e3 = sub
                substr1 = page1.extract_text().split(' \n \n')
                for sub in substr1:
                    if not sub.find('GR 1') and not sub.find('SZTUCZNA') < 0 and not sub.isspace() and sub.find('godz'):
                        for _ in range(2):
                            sub = sub.removeprefix('\n')
                            sub = sub.removeprefix(' \n')
                            sub = sub.removeprefix(' ')
                        if sub.find('8.15') > 0: c_b4 = sub
                        elif sub.find('10.45') > 0: c_c4 = sub
                        elif sub.find('13.45') > 0: c_d4 = sub
                        elif sub.find('16.15') > 0: c_e4 = sub
                format_excel(workbook, c_b3, c_c3, c_d3, c_e3, c_b4, c_c4, c_d4, c_e4)
                break
workbook.close()
writer.write(export + '/Schedule.pdf')



