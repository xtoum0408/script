######################################
# 点位表中有位置名称重复，需添加编号加以区分
#######################################


import openpyxl as xl

book = xl.load_workbook('C:/Users/CSY/Desktop/门禁管理系统点位表.xlsx')
sheet = book['办公门禁']

num = 1
for row in range(4, 589, 1):
    sheet[f'B{row}'].value = 'R' + str(num).zfill(3)
    num += 1

    value = sheet[f'H{row}'].value
    if value is not None:
        luyou = sheet[f'H{row}'].value
        sheet[f'H{row}'].value = sheet[f'B{row}'].value + luyou[7:]

merged_cells = sheet.merged_cells
mc = []
for merged_area in merged_cells:
    r1, r2 = merged_area.min_row, merged_area.max_row
    if r2 - r1 > 0 and 'C' in str(merged_area):
        mc.append((r1, r2))
        # print(merged_area)
t = []

for m in mc:
    first_row = m[0]
    last_row = m[1] + 1

    for row in range(first_row, last_row):
        n = 2
        first = row + 1
        area = sheet[f'D{row}'].value
        for r in range(first, last_row):
            tar = sheet[f'D{r}'].value
            if tar == area:
                if n == 2:
                    sheet[f'D{row}'].value = sheet[f'D{row}'].value + '-1'
                # print('旧', sheet[f'D{r}'].value, r)
                sheet[f'D{r}'].value = sheet[f'D{r}'].value + '-' + str(n)
                # print(sheet[f'D{r}'].value)
                n += 1



book.save('test.xlsx')
print('-----------------完成-----------------')
