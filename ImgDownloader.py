import os
import urllib.request
import xlrd

headers = {
    "Cookie": "UM_distinctid=16685e0279d3e0-06f34603dfa898-36664c08-1fa400-16685e0279e133; bdshare_firstime=1539844405694; gsScrollPos-1702681410=; CNZZDATA1254092508=1744643453-1539842703-%7C1539929860; _d_id=0ba0365838c8f6569af46a1e638d05",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36",
}
# 确定图片保存路径
path = "images/"
if not os.path.exists(path):
    os.mkdir(path)
# 打开excel文件
excelpath = input("请将excel文件拖入此窗口，然后回车继续\n")
data = xlrd.open_workbook(excelpath)
table = data.sheet_by_name("Sheet0")

# 确定需要下载的列，可以是多个列
allx = table.nrows  # 总行数
namex = table.row_values(0)  # 所有的题头
setxs = []
set_name = []
for i in range(len(namex)):
    print('[' + str(i) + ']', namex[i])
print('请输入所需下载的列的数字（如需多个列用空格隔开）')
tl = list(map(eval, input().split()))
for i in tl:
    setxs.append(i)
    set_name.append(namex[i])
    if not os.path.exists(path + namex[i]):
        os.mkdir(path + namex[i])

# 开始遍历下载
# 先遍历指定下载的列，再遍历每一列中的单元格内容
# 如果单元格内容有多个链接，转化为列表后再依次下载
for i in range(len(setxs)):
    setx = setxs[i]
    t = 1
    while t < allx:
        name = table.cell_value(t, 0)  # 提交人名字
        srclist = table.cell_value(t, setx).split(',')  # 图片链接
        t = t + 1
        print(name)
        print(srclist)
        for i in range(len(srclist)):
            src = srclist[i]
            urllib.request.urlretrieve(src, path + set_name[i] + '/' + name + '-' + str(i + 1) + ".jpg")
            print("-------- downloading ---------")
            print("------ download done -------")
