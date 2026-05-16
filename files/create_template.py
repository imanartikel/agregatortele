import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Upload"

headers = [
    "No", "Jenis Pengiriman", "Pilihan Kurir", "Jenis Layanan",
    "Nama Pengirim", "Alamat Pengirim", "No Telp Pengirim", "Kecamatan Pengirim", "Kota Pengirim",
    "Nama Penerima", "Alamat Penerima", "No Telp Penerima", "Kecamatan Penerima", "Kota Penerima",
    "Berat", "Panjang", "Lebar", "Tinggi", "Isi Paket", "Asuransi",
    "Harga Barang", "Nilai COD", "Instruksi", "No Referensi"
]

for col, header in enumerate(headers, 1):
    ws.cell(row=1, column=col, value=header)

wb.save("template.xlsx")
print("template.xlsx created!")
