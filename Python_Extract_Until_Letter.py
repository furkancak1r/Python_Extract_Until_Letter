# Telefon numaralarındaki harflere kadar olan kısmı alan kod
import openpyxl
import re  # Regular expression kütüphanesini ekleyin

def extract_until_letter(cell_value):
    # İlk harfe kadar olan kısmı bulmak için regular expression kullanın
    match = re.search("[a-zA-Z]", cell_value)
    if match:
        return cell_value[:match.start()]
    else:
        return cell_value

# Excel dosyasını aç
wb = openpyxl.load_workbook(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")

sheet = wb['INKOOL']  # Sayfa adını buraya yazın

# J5'ten J7583'e kadar olan hücreleri dolaş
for row in range(5, 7584):
    cell_value = sheet[f'H{row}'].value

    # Eğer hücre boş değilse, harfe kadar olan kısmı çıkar ve L sütununa yaz
    if cell_value is not None:
        extracted_value = extract_until_letter(cell_value)
        small_letters = extracted_value.lower()
        sheet[f'L{row}'].value = small_letters

# Excel dosyasını kaydet
wb.save(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")
