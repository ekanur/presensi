import requests, os, sys, csv, calendar, datetime

from bs4 import BeautifulSoup
from pyhtml2pdf import converter

daftar_bulan    = ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember"]
bulan           = sys.argv[1] if len(sys.argv) > 1 else daftar_bulan[datetime.date.today().month-1]
login_url       = "https://simpeg2.jogjaprov.go.id/prod/index.php/login/cek_login"
logout_url      = "https://simpeg2.jogjaprov.go.id/prod/index.php/login/logout"
scrap_page      = "https://simpeg2.jogjaprov.go.id/prod/index.php/lap_pres_harian/cetak/"

session         = requests.session()

angka_bulan     = daftar_bulan.index(bulan)+1 if bulan != 'desember' else 12
bulan_ttd       = daftar_bulan[angka_bulan] if angka_bulan < 12 else daftar_bulan[0]
tahun           = datetime.date.today().year

if os.path.exists("rekap") == False : os.mkdir("rekap")
if os.path.exists("pdf") == False : os.mkdir("pdf")

def read_data(file):
    data = []
    
    with open(file) as csv_file:
        reader = csv.reader(csv_file, delimiter=',')
        for row in reader:
            data.append({"nip":row[0], "password":row[1], 'nama':row[2]})
    
    return data

def get_presensi(akun):

    with session as s: 
        payload = {  
            "username": akun['nip'], 
            "password": akun['password']
        } 
        
        cek_login = s.post(login_url, data=payload)
        
        is_loggedin = BeautifulSoup(cek_login.content, "html.parser")

        if(len(is_loggedin.select("div.alert-error"))>0):
            return []

        # https://simpeg2.jogjaprov.go.id/prod/index.php/lap_pres_harian/cetak/01-08-2023/31-08-2023
        awal_bulan = "01-"+str((daftar_bulan.index(bulan)+1))+"-"+str(tahun)
        akhir_bulan = str(calendar.monthrange(tahun, (daftar_bulan.index(bulan)+1))[1])+"-"+str((daftar_bulan.index(bulan)+1))+"-"+str(tahun)
        r = s.get(scrap_page+awal_bulan+"/"+akhir_bulan)
        soup = BeautifulSoup (r.content, "html.parser")
        page_title = soup.select("div.title_left>table")
        presensi = soup.select("div.x_content")


        ttd = create_sign(akun['nama'], akun['nip'], bulan_ttd)
        s.get(logout_url)

        return [page_title[0], presensi[0], ttd]

def create_sign(nama_lengap, nip, bulan_ttd):
    tahun_ttd = tahun if bulan_ttd != 'januari' else tahun+1
    img_ttd = f"<img src='../ttd/{nip}.jpg' width='50px' height='58px'/><br/>" if os.path.exists(f'ttd/{nip}.jpg') else '<br/><br/><br/><br/>'
    content = f'''
    <table border=0 width='100%'>
        <tr>
            <td width='70%'>
                Mengetahui,<br/>
                Kepala Sekolah<br/><br/><br/><br/><br/><br/>
                <strong>
                    <u>Drs. Agus Waluyo, M.Eng</u><br/>
                    196512271994121002
                </strong>
            </td>
            <td width='30%'>
            Sleman,&nbsp; &nbsp &nbsp; {bulan_ttd.capitalize()} {tahun_ttd}<br/><br/>
            Guru<br/>
            {img_ttd}
            <strong>
                <u>{nama_lengap}</u><br/>
                {nip}
            </strong>
            </td>
        </tr>
    </table>
    '''
    return content

def write_presensi() :
    data = read_data("akun.csv")
    for akun in data :

        content = get_presensi(akun)

        if(content != []):
            rekap = BeautifulSoup(f'''
                    <html>
                        <head>
                        <title>Presensi Bulanan</title>
                        <link rel='stylesheet' href='../style.css'/>
                        
                        </head>
                        <body>
                        {content[0]}
                        {content[1]}
                        {content[2]}
                        </body>
                    </html>
                    ''', 'html.parser')
            with open("rekap/"+akun["nama"]+".html", 'w', encoding='utf-8') as file:

                file.write(rekap.prettify())

            converter.convert(os.path.abspath(f'rekap/{akun["nama"]}.html'), f"pdf/{akun['nama']}_{daftar_bulan[angka_bulan]}.pdf", print_options={"paperHeight":13, "paperWidth":8.27})
            os.remove(f"rekap/{akun['nama']}.html")

            print(f"Berhasil membuat rekap {akun['nama']}_{daftar_bulan[angka_bulan]}.pdf di folder pdf")
        else:
            print(f"Gagal membuat rekap {akun['nama']}. Akun gagal login")


# print(type(tahun))
write_presensi()

