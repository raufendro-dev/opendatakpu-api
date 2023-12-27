import openpyxl
import json
def dalamnegeri(query):
    workbook = openpyxl.load_workbook("dalamnegeri.xlsx")

    worksheet = workbook.active  # Or access a specific sheet by name: worksheet = workbook["Sheet1"]

    # Choose the column you want to extract (e.g., column B)
    kota_extract = "B"
    jumlahkecamatan_extract = "C"
    jumlahkelurahan_extract = "D"
    jumlahtps_extract = "E"
    lakilaki_extract = "F"
    perempuan_extract = "G"
    jumlah_extract = "H"
    

    # Get all values in the specified column
    kota_values = [cell.value for cell in worksheet[kota_extract]]
    jumlahkecamatan_values = [cell.value for cell in worksheet[jumlahkecamatan_extract]]
    jumlahkelurahan_values = [cell.value for cell in worksheet[jumlahkelurahan_extract]]
    jumlahtps_values = [cell.value for cell in worksheet[jumlahtps_extract]]
    lakilaki_values = [cell.value for cell in worksheet[lakilaki_extract]]
    perempuan_values = [cell.value for cell in worksheet[perempuan_extract]]
    jumlah_values = [cell.value for cell in worksheet[jumlah_extract]]

    if query == "":
            kota = []
            kota = kota_values
            kota.remove('Nama Kabupaten/Kota')
            kota.remove(None)
            kota.remove(None)

            jumlahkecamatan = []
            jumlahkecamatan = jumlahkecamatan_values
            jumlahkecamatan.remove('Jumlah Kecamatan')
            jumlahkecamatan.remove(None)
            jumlahkecamatan.remove(None)

            jumlahkelurahan = []
            jumlahkelurahan = jumlahkelurahan_values
            jumlahkelurahan.remove('Jumlah Kelurahan/Desa')
            jumlahkelurahan.remove(None)
            jumlahkelurahan.remove(None)

            jumlahtps = []
            jumlahtps = jumlahtps_values
            jumlahtps.remove('Jumlah TPS')
            jumlahtps.remove(None)
            jumlahtps.remove(None)

            lakilaki = []
            lakilaki = lakilaki_values
            lakilaki.remove('Total Penetapan Pemilih DPT')
            lakilaki.remove('Laki-Laki')
            lakilaki.remove(None)

            perempuan = []
            perempuan = perempuan_values
            perempuan.remove(None)
            perempuan.remove('Perempuan')
            perempuan.remove(None)


            jumlah = []
            jumlah = jumlah_values
            jumlah.remove(None)
            jumlah.remove('Jumlah')
            jumlah.remove(None)

            data = []

            for i in range(len(kota)):
                if kota[i]!= None:
                    value = {
                        'kota':kota[i],
                        'jumlahkecamatan':jumlahkecamatan[i],
                        'jumlahkelurahan':jumlahkelurahan[i],
                        'jumlahtps':jumlahtps[i],
                        'lakilaki':lakilaki[i],
                        'perempuan':perempuan[i],
                        'jumlah':jumlah[i]
                    }
                    data.append(value)
   
            hasil = json.dumps(data)
           
            return hasil
    else :
        id = 0
        found = 0
        try:
            index = kota_values.index(query.upper())
            id = index
            found = 1
        except ValueError:
            found = 0
        
        if found != 0:
            value={
                'kota':kota_values[id],
                'jumlahkecamatan':jumlahkecamatan_values[id],
                'jumlahkelurahan':jumlahkelurahan_values[id],
                'jumlahtps':jumlahtps_values[id],
                'lakilaki':lakilaki_values[id],
                'perempuan':perempuan_values[id],
                'jumlah':jumlah_values[id]
            }
            hasil = json.dumps(value)
    
            return hasil
        else:
             value = {
                  "error":True,
                  "pesan":"Daerah tidak ditemukan"
             }
             hasil = json.dumps(value)
             return hasil



def profilpartai(query):
    workbook = openpyxl.load_workbook("profilpartai.xlsx")

    worksheet = workbook.active  # Or access a specific sheet by name: worksheet = workbook["Sheet1"]

    # Choose the column you want to extract (e.g., column B)
    namaparpol_extract = "B"
    akronim_extract = "C"
    nosklambangpp_extract = "D"
    nobnppterdaftar_extract = "E"
    tglbnpp_extract = "F"
    namanotaris_extract = "G"
    noaktenotaris_extract = "H"
    alamatnotaris_extract = "I"
    tglaktanotaris_extract = "J"
    noadart_extract = "K"
    tgladart_extract = "L"
    email_extract = "M"
    website_extract = "N"
    nourutpemilu2024_extract = "O"


    

    # Get all values in the specified column
    namaparpol_values = [cell.value.upper() for cell in worksheet[namaparpol_extract]]
    akronim_values = [cell.value.upper() for cell in worksheet[akronim_extract]]
    nosklambangpp_values = [cell.value for cell in worksheet[nosklambangpp_extract]]
    nobnppterdaftar_values = [cell.value for cell in worksheet[nobnppterdaftar_extract]]
    nobnppterdaftar_values = [cell.value for cell in worksheet[nobnppterdaftar_extract]]
    tglbnpp_values = [cell.value for cell in worksheet[tglbnpp_extract]]
    namanotaris_values = [cell.value for cell in worksheet[namanotaris_extract]]
    namanotaris_values = [cell.value for cell in worksheet[namanotaris_extract]]
    noaktenotaris_values = [cell.value for cell in worksheet[noaktenotaris_extract]]
    alamatnotaris_values = [cell.value for cell in worksheet[alamatnotaris_extract]]
    tglaktanotaris_values = [cell.value for cell in worksheet[tglaktanotaris_extract]]
    noadart_values = [cell.value for cell in worksheet[noadart_extract]]
    tgladart_values = [cell.value for cell in worksheet[tgladart_extract]]
    email_values = [cell.value for cell in worksheet[email_extract]]
    website_values = [cell.value for cell in worksheet[website_extract]]
    nourutpemilu2024_values = [cell.value for cell in worksheet[nourutpemilu2024_extract]]


    if query == "":
            namaparpol = []
            namaparpol = namaparpol_values


            akronim = []
            akronim = akronim_values

            nosklambangpp = []
            nosklambangpp = nosklambangpp_values

            nobnppterdaftar = []
            nobnppterdaftar = nobnppterdaftar_values

            tglbnpp = []
            tglbnpp = tglbnpp_values

            namanotaris = []
            namanotaris = namanotaris_values

            noaktenotaris = []
            noaktenotaris = noaktenotaris_values
            
            alamatnotaris = []
            alamatnotaris = alamatnotaris_values

            tglaktanotaris = []
            tglaktanotaris = tglaktanotaris_values

            noadart = []
            noadart = noadart_values

            tgladart = []
            tgladart = tgladart_values

            email =[]
            email = email_values
            
            website =[]
            website = website_values

            nourutpemilu2024 = []
            nourutpemilu2024 = nourutpemilu2024_values




            data = []

            for i in range(len(namaparpol)):
                if namaparpol[i]!= None:
                    value = {
                        'namaparpol':str(namaparpol[i]),
                        'akronim':str(akronim[i]),
                        'no_sk_lambang_pp':str(nosklambangpp[i]),
                        'no_bn_pp_terdaftar': str(nobnppterdaftar[i]),
                        'tgl_bn_pp': str(tglbnpp[i]),
                        'nama_notaris': str(namanotaris[i]),
                        'no_akte_notaris':str(noaktenotaris[i]),
                        'alamat_notaris':str(alamatnotaris[i]),
                        'tgl_akta_notaris':str(tglaktanotaris[i]),
                        'no_ad_art': str(noadart[i]),
                        'tgl_ad_art':str(tgladart[i]),
                        'email': str(email[i]),
                        'website':str(website[i]),
                        'no_urut_pemilu_2024':str(nourutpemilu2024[i])


                    }
                    data.append(value)
   
            hasil = json.dumps(data)
            print(hasil)
           
            return hasil
    else :
        id = 0
        found = 0
        try:
            index = akronim_values.index(query.upper())
            id = index
            found = 1
        except ValueError:
            found = 0
        
        if found != 0:
            value={
                'namaparpol':str(namaparpol_values[id]),
                'akronim':str(akronim_values[id]),
                'no_sk_lambang_pp':str(nosklambangpp_values[id]),
                'no_bn_pp_terdaftar':str(nobnppterdaftar_values[id]),
                'tgl_bn_pp': str(tglbnpp_values[id]),
                'nama_notaris': str(namanotaris_values[id]),
                'no_akte_notaris':str(noaktenotaris_values[id]),
                'alamat_notaris':str(alamatnotaris_values[id]),
                'tgl_akta_notaris':str(tglaktanotaris_values[id]),
                'no_ad_art': str(noadart_values[id]),
                'tgl_ad_art':str(tgladart_values[id]),
                'email': str(email_values[id]),
                'website':str(website_values[id]),
                'no_urut_pemilu_2024':str(nourutpemilu2024_values[id])
                

            }
            hasil = json.dumps(value)
            print(hasil)
            return hasil
        else:
             value = {
                  "error":True,
                  "pesan":"Partai Politik tidak ditemukan"
             }
             hasil = json.dumps(value)
             print(hasil)
             return hasil
        


