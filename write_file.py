from PIL import Image as Images
from PIL import ImageFont
from PIL import ImageDraw

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

def write_img(data):
    with Images.open("rasm/maket.jpg").convert("RGBA") as base:
        txt = Images.new("RGBA", base.size, (0, 0, 0, 0))

        d = ImageDraw.Draw(txt)

        fnt = ImageFont.truetype("shrift/Roboto-Regular.ttf", 8)

        # draw text, half opacity
        d.text((127, 96), f"{data['Loyiha_nomi']}", font=fnt, fill=(0, 0, 0))
        d.text((127, 106), f"{data['Sana']}", font=fnt, fill=(0, 0, 0))
        d.text((127, 116), f"{data['Menedjer']}", font=fnt, fill=(0, 0, 0))
        if data['Shahta_turi'] == "Beton":
            d.text((125, 126), "+", font=fnt, fill=(0, 0, 0))
        elif data['Shahta_turi'] == "Bloklar":
            d.text((184, 126), "+", font=fnt, fill=(0, 0, 0))
        else:
            d.text((238, 126), "+", font=fnt, fill=(0, 0, 0))

        if data['Mrl_Or_Rl'] == "MR":
            d.text((173, 146), "+", font=fnt, fill=(0, 0, 0))
        else:
            d.text((218, 146), "+", font=fnt, fill=(0, 0, 0))

        d.text((313, 190), f"{data['olchamlar']['A']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 200), f"{data['olchamlar']['B']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 210), f"{data['olchamlar']['C']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 220), f"{data['olchamlar']['D']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 230), f"{data['olchamlar']['E']}", font=fnt, fill=(0, 0, 0))

        i = 292
        for item in data['qavatlar']:
            d.text((313, i), f"{data['qavatlar'][f'{item}']}", font=fnt, fill=(0, 0, 0))
            i += 10

        d.text((313, 455), f"{len(data['qavatlar'])}", font=fnt, fill=(0, 0, 0))
        d.text((313, 465), f"{data['eshiklar_soni']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 475), f"{data['P']}", font=fnt, fill=(0, 0, 0))
        d.text((313, 485), f"{data['Hp']}", font=fnt, fill=(0, 0, 0))
        d.text((261, 516), f"{data['Eslatma']}", font=fnt, fill=(0, 0, 0))

        out = Images.alpha_composite(base, txt)

        out.save("salom.png")

def write_excel_osten(keys, data):
    wb = Workbook()
    dest_filename = f"{data['Loyiha_nomi']}.xlsx"
    ws1 = wb.active
    ws1.title = "Olchamlar"
    ws1.sheet_properties.tabColor = "1072BA"

    ws1.column_dimensions["A"].width = 20

    ws1.column_dimensions["C"].width = 20
    ws1.column_dimensions["D"].width = 25
    ws1.column_dimensions["E"].width = 30
    ws1.column_dimensions["F"].width = 20
    ws1.column_dimensions["G"].width = 20

    # logo
    ws1.cell(row=1, column=1).border = thin_border
    img = Image('rasm/logo.png')
    ws1.merge_cells('A1:B5')
    ws1.add_image(img, 'A1')

    cell = ws1.cell(row=1, column=3)
    cell.value = "O'lcham loyihalari"
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=1, column=3).border = thin_border
    ws1.merge_cells('C1:E5')

    for i in range(7, 10):
        ws1.cell(row=i, column=1).border = thin_border
        ws1.cell(row=i, column=2).border = thin_border

    ws1.cell(row=7, column=1, value=keys[0])
    ws1.merge_cells('B7:E7')
    ws1.cell(row=7, column=2, value=data['Loyiha_nomi'])

    ws1.cell(row=8, column=1, value=keys[1])
    ws1.merge_cells('B8:E8')
    ws1.cell(row=8, column=2, value=data[f'{keys[1]}'])

    ws1.cell(row=9, column=1, value=keys[2])
    ws1.merge_cells('B9:E9')
    ws1.cell(row=9, column=2, value=data[f'{keys[2]}'])

    ws1.cell(row=10, column=1, value=keys[3])
    shahta = ["Beton", "Bloklar", "Po'lat Konstruksiya"]
    if data['Shahta_turi'] == "Beton":
        text = f"+ {data['Shahta_turi']}"
        ws1.cell(row=10, column=3, value=text).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=4, value=shahta[1]).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=5, value=shahta[2]).alignment = Alignment(horizontal='center', vertical='center')
    elif data['Shahta_turi'] == "Bloklar":
        text = f"+ {data['Shahta_turi']}"
        ws1.cell(row=10, column=4, value=text).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=3, value=shahta[0]).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=5, value=shahta[2]).alignment = Alignment(horizontal='center', vertical='center')
    else:
        text = f"+ {data['Shahta_turi']}"
        ws1.cell(row=10, column=5, value=text).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=3, value=shahta[0]).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=10, column=4, value=shahta[1]).alignment = Alignment(horizontal='center', vertical='center')

    for i in range(3, 6):
        ws1.cell(row=10, column=i).border = thin_border

    MRL_MR = ["MR", "MRL"]
    if data['Mrl_Or_Rl'] == "MR":
        text = f"+ {data['Mrl_Or_Rl']}"
        ws1.cell(row=11, column=1, value=text).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=11, column=4, value=MRL_MR[1]).alignment = Alignment(horizontal='center', vertical='center')
    elif data['Mrl_Or_Rl'] == "MRL":
        text = f"+ {data['Mrl_Or_Rl']}"
        ws1.cell(row=11, column=1, value=MRL_MR[0]).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=11, column=4, value=text).alignment = Alignment(horizontal='center', vertical='center')

    ws1.cell(row=11, column=1).border = thin_border
    ws1.cell(row=11, column=4).border = thin_border

    ws1.merge_cells('A11:C11')
    ws1.merge_cells('D11:E11')

    img = Image("rasm/rasm_1.jpg")
    img.width = 250
    img.height = 240
    ws1.add_image(img, 'A13')
    img = Image("rasm/rasm_2.jpg")
    img.width = 250
    img.height = 430
    ws1.add_image(img, 'A27')
    ws1.merge_cells('A12:C51')

    text = "olcham"
    ws1.cell(row=12, column=5, value=text).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=12, column=4).border = thin_border
    ws1.cell(row=12, column=5).border = thin_border

    olcham = []
    for item in data['olchamlar']:
        olcham.append(item)
    ws1.cell(row=13, column=4, value=olcham[0]).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=13, column=5, value=data["olchamlar"]["A"]).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')

    ws1.cell(row=14, column=4, value=olcham[1]).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=14, column=5, value=data["olchamlar"]["B"]).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')

    ws1.cell(row=15, column=4, value=olcham[2]).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=15, column=5, value=data["olchamlar"]["C"]).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')

    ws1.cell(row=16, column=4, value=olcham[3]).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=16, column=5, value=data["olchamlar"]["D"]).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')

    ws1.cell(row=17, column=4, value=olcham[4]).alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=17, column=5, value=data["olchamlar"]["E"]).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')
    for i in range(13, 18):
        ws1.cell(row=i, column=4).border = thin_border
        ws1.cell(row=i, column=5).border = thin_border

    ws1.cell(row=20, column=4, value=f'Nomer').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=20, column=5, value=f'Balandlik(H)').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=20, column=4).border = thin_border
    ws1.cell(row=20, column=5).border = thin_border
    i = 21
    for item, value in data['qavatlar'].items():
        ws1.cell(row=i, column=4, value=item).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=i, column=5, value=value).alignment = Alignment(horizontal='center', vertical='center')
        ws1.cell(row=i, column=4).border = thin_border
        ws1.cell(row=i, column=5).border = thin_border
        i += 1

    ws1.cell(row=37, column=4, value=f'Qavatlar soni').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=37, column=5, value=len(data['qavatlar'])).alignment = Alignment(horizontal='center',
                                                                                  vertical='center')

    ws1.cell(row=38, column=4, value=f'Eshiklar soni').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=38, column=5, value=data['eshiklar_soni']).alignment = Alignment(horizontal='center',
                                                                                  vertical='center')

    ws1.cell(row=39, column=4, value=f'Priyamka balandligi').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=39, column=5, value=data['P']).alignment = Alignment(horizontal='center', vertical='center')

    ws1.cell(row=40, column=4, value=f'Oxirgi etaj balandligi').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=40, column=5, value=data['Hp']).alignment = Alignment(horizontal='center', vertical='center')

    ws1.cell(row=41, column=4, value=f'Eslatma').alignment = Alignment(horizontal='center', vertical='center')
    ws1.cell(row=41, column=5, value=data['Eslatma']).alignment = Alignment(horizontal='justify', vertical='justify')

    for i in range(37, 42):
        ws1.cell(row=i, column=4).border = thin_border
        ws1.cell(row=i, column=5).border = thin_border
    ws1.merge_cells('E41:E49')
    wb.save(filename=dest_filename)
    print("ishladi")

def write_excel_birnima(data):
    wb = Workbook()
    dest_filename = f"{data['Лойиҳа номи']}.xlsx"
    ws2 = wb.active
    ws2.title = "Olchamlar"
    ws2.sheet_properties.tabColor = "1072BA"

    # ws2.column_dimensions["A"].width = 15
    # ws2.column_dimensions["B"].width = 15
    # ws2.column_dimensions["C"].width = 15
    # ws2.column_dimensions["D"].width = 15

    # logo
    img = Image('rasm/rasm_2_bot.png')
    ws2.merge_cells('H2:N22')
    ws2.add_image(img, 'H2')

    i = 2
    for item in data:
        if item != 'state':
            ws2.cell(row=i, column=1, value=f'{item}:')
            ws2.cell(row=i, column=4, value=f"{data[f'{item}']}").alignment = Alignment(horizontal='center',vertical='center')
            ws2.cell(row=i, column=1).border = thin_border
            ws2.cell(row=i, column=4).border = thin_border

            ws2.merge_cells(f'A{i}:C{i}')
            ws2.merge_cells(f'D{i}:F{i}')
            i += 1
        else:
            break
    qavat_keys = ["A", "B", "C", "D", "E", "F"]
    d = 1
    e = 2
    for j in range(1, 10):
        if len(data[f'qavat_{j}']) != 0:
            ws2.cell(row=24, column=d, value=f'№')
            ws2.cell(row=24, column=e, value=f"Этаж - {j}").alignment = Alignment(horizontal='center',vertical='center')
            ws2.cell(row=24, column=d).border = thin_border
            ws2.cell(row=24, column=e).border = thin_border

            for i in range(1, 7):
                ws2.cell(row=24 + i, column=d, value=f'{qavat_keys[i - 1]}')
                son = data[f'qavat_{j}'][qavat_keys[i - 1]]
                ws2.cell(row=24 + i, column=e, value=son).alignment = Alignment(horizontal='center', vertical='center')
                ws2.cell(row=24 + i, column=d).border = thin_border
                ws2.cell(row=24 + i, column=e).border = thin_border

        d += 2
        e += 2

    d = 1
    e = 2
    for j in range(10, 19):
        if len(data[f'qavat_{j}']) != 0:
            ws2.cell(row=32, column=d, value=f'№')
            ws2.cell(row=32, column=e, value=f"Этаж - {j}").alignment = Alignment(horizontal='center',vertical='center')
            ws2.cell(row=32, column=d).border = thin_border
            ws2.cell(row=32, column=e).border = thin_border

            for i in range(1, 7):
                ws2.cell(row=32 + i, column=d, value=f'{qavat_keys[i - 1]}')
                son = data[f'qavat_{j}'][qavat_keys[i - 1]]
                ws2.cell(row=32 + i, column=e, value=son).alignment = Alignment(horizontal='center', vertical='center')
                ws2.cell(row=32 + i, column=d).border = thin_border
                ws2.cell(row=32 + i, column=e).border = thin_border
        d += 2
        e += 2
    wb.save(filename=dest_filename)