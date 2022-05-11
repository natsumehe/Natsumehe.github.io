import xlwings as xw

app = xw.App(visible = True,add_book = False)
app.display_alerts = False

wb = xw.Book('E:\code\data.csv\扶贫.xlsx')

sheet = wb.sheets['入账主体的地点']

feixi = sheet.range('A1:A73').value
changfeng = sheet .range('C1:C58').value
feidong = sheet.range('E1:E66').value
lujiang = sheet.range('G1:G94').value
chaohu = sheet.range('I1:I128').value

a = feixi + changfeng + feidong + lujiang + chaohu 
# print(a,len(a))
print(len(feixi),len(changfeng),len(feidong),len(lujiang),len(chaohu))

sum = len(feixi) + len(changfeng) + len(feidong) + len(lujiang) + len(chaohu)
print(sum)

print(feidong)
print(len(feidong ))
wb.close
app.quit