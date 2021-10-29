import requests, xlwt, os, pyfiglet, time, sys
from bs4 import BeautifulSoup


def clear():
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


clear()

print(pyfiglet.figlet_format('ULSMAN', font='slant'))
print('Gunakan dengan bijak jangan sampai guru mengetahui ini\nAuthor: Nizar\n')

url_base = 'https://sman1kaliwungu.sch.id/elearning/ujian/'

ulangan = input('Nama Ulangan: ')
id_ulangan = int(input('Id Ulangan: '))
print()

try:
    count_url = f'{url_base}muncul.php?kirim={str(id_ulangan)}&nama={str(ulangan)}'

    page = requests.get(count_url)
    soup = BeautifulSoup(page.content, 'html.parser')
    table = soup.find('tbody')
    banyak_soal = len(table.find_all('tr'))

    start = 1
    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet('Sheet 1')

    aligment = xlwt.Alignment()
    aligment.vert = aligment.VERT_TOP
    aligment.wrap = aligment.WRAP_AT_RIGHT

    style = xlwt.XFStyle()
    style.alignment = aligment

    sheet1.col(1).width = 10120
    sheet1.col(2).width = 5120
    sheet1.col(3).width = 5120
    sheet1.col(4).width = 5120
    sheet1.col(5).width = 5120
    sheet1.col(6).width = 5120

    sheet1.write(0, 0, 'No')
    sheet1.write(0, 1, 'Soal')
    sheet1.write(0, 2, 'A')
    sheet1.write(0, 3, 'B')
    sheet1.write(0, 4, 'C')
    sheet1.write(0, 5, 'D')
    sheet1.write(0, 6, 'E')
    sheet1.write(0, 7, 'Kunci')

    while (start <= banyak_soal):
        url = f'{url_base}buat.php?no={str(start)}&id={str(id_ulangan)}'
        page_soal = requests.get(url)
        soup_soal = BeautifulSoup(page_soal.content, 'html.parser')
        form = soup_soal.find('form')

        # get teks
        j1 = form.find('textarea', {'name': 'j1'})
        ja = form.find('textarea', {'name': 'ja'})
        jb = form.find('textarea', {'name': 'jb'})
        jc = form.find('textarea', {'name': 'jc'})
        jd = form.find('textarea', {'name': 'jd'})
        je = form.find('textarea', {'name': 'je'})
        jkunci = form.find('input', {'name': 'kunci'}).get('value')

        soal = j1.text
        a = ja.text
        b = jb.text
        c = jc.text
        d = jd.text
        e = je.text
        kunci = jkunci

        sheet1.write(start, 0, int(start), style)

        if j1.find('img'):
            gambar = j1.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 1, f'{soal}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 1, soal, style)

        if ja.find('img'):
            gambar = ja.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 2, f'{a}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 2, a, style)

        if jb.find('img'):
            gambar = jb.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 3, f'{b}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 3, b, style)

        if jc.find('img'):
            gambar = jc.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 4, f'{c}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 4, c, style)

        if jd.find('img'):
            gambar = jd.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 5, f'{d}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 5, d, style)

        if je.find('img'):
            gambar = je.find('img')['src']
            url_gambar = url_base + gambar
            sheet1.write(start, 6, f'{e}\n\nAda gambar, link:\n{url_gambar}', style)
        else:
            sheet1.write(start, 6, e, style)

        sheet1.write(start, 7, kunci, style)

        print('[+] Berhasil Ambil Soal ' + str(start))

        start += 1

    time.sleep(5)
    clear()
    print(pyfiglet.figlet_format('ULSMAN', font='slant'))
    print('Gunakan dengan bijak jangan sampai guru mengetahui ini\n')
    print(f'[+] Berhasil Ambil Soal Ulangan {ulangan}')
    book.save(f'{ulangan}-ulangan.xls')
except:
    print('[-] Ulangan tidak ada')
    print('[-] Ulangi dan masukkan input dengan benar')
