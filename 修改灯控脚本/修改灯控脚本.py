######################################
# 灯控脚本有十二万行，需批量修改其中参数
######################################

lines = []
c = 2
index = 0
data_point_list = []
target = 'Read="True" Receive="True" Write="True"'
unrename = 'Receive="True"'

with open('单RWI.txt', 'r', encoding='utf-8') as file:
    data = file.readlines()

for L in data:
    lines.append(L)

for l in lines:
    if l.count('</Datapoint>') > 0:
        data_point_list.append(c)   # 通过</Datapoint>加两行定位到需要修改的位置，即lines这个list中的索引值
    c += 1

# print(data_point_list)
# print(lines[data_point_list[20]])


for xiugai in data_point_list[:-1]:
    # print(lines[xiugai])
    if lines[xiugai].count(target) > 0:
        continue
    else:
        print('喜加一')
        lines[xiugai] = lines[xiugai].replace(unrename, target)
        print(lines[xiugai])


with open('text.txt', 'w') as t:
    for line_input in lines:
        t.write(line_input)

print('完成')


