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





