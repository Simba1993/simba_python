#********************读取一个excel表的内容并存储在3个列表中********************
import xlrd
import xlwt
#srdz=input("输入你的输入文件路径:")
#scdz=input("输入你的输出文件路径:")
data = xlrd.open_workbook(r"C:\Users\scyantao\Desktop\test1\2.xlsx")
# data.sheet_names()
sheet = data.sheet_by_index(0)
harddisk1 = []
harddisk2 = []
harddisk3 = []
harddisk4 = []
# print("sheet_name:",data.sheet_names())
# print("sheet_number:",data.nsheets)
hangshu = sheet.nrows  # 确认sheet1的行数

#判断标号在第几列
firsthang= (sheet.row_values(0))
for i in range(0,len(firsthang)):
    if firsthang[i]=="标号":
        num=i

        harddisk1 = (sheet.col_values(num))  # 将“标号”列的数放到harddisk1里面
        # print(harddisk1)
        for i in range(1, len(harddisk1)):
            if harddisk1[i] == '':
                harddisk1[i] = harddisk1[i - 1]
        del harddisk1[0]  # 去除无用的第一个数

        harddisk2 = (sheet.col_values(num+2))  #将“电缆长度米”列的米数放到harddisk2里面
        for i in range(1, len(harddisk1)):
            if harddisk2[i] == '':
                harddisk2[i] = harddisk2[i - 1]
        del harddisk2[0]  # 去除无用的第一个数


        harddisk3 = (sheet.col_values(num+3)) #将“配置说明”列的米数放到harddisk3里面
        del harddisk3[0]

        #**************让harddisk3的数据输出在“ ”之间*****************
        datas=harddisk3
        harddisk4=[]
        for mm in datas:
                flag=0
                flag1=0
                for i in mm:
                    flag=flag+1
                    if i=='“':
                        break
                for i in mm:
                    flag1=flag1+1
                    if i=='”':
                        break
                ss=mm[flag:flag1-1]
                harddisk4.append(ss)
    #**************************************************************

    #****************将数据按固定的格式写入EXCEL*********************
        work_book = xlwt.Workbook(encoding = 'ascii')
        work_sheet = work_book.add_sheet('My Worksheet')
        style1 = xlwt.XFStyle() # 初始化样式
        font = xlwt.Font() # 为样式创建字体
        font.name='宋体'
        font.colour_index = 0  #设置字体颜色：黑色0、红色2，白色1、蓝色4、黄色5
        font.NSimSun=True
        font.height= 11*20  #设置字体大小为11号字体
        style1.font=font

        #给单元格加框线
        border = xlwt.Borders()
        border.left = xlwt.Borders.THIN  #左
        border.top=xlwt.Borders.THIN     #上
        border.right=xlwt.Borders.THIN   #右
        border.bottom=xlwt.Borders.THIN  #下
        border.left_colour = 0x40  #设置框线颜色，0x40是黑色
        border.right_colour = 0x40 #设置框线颜色，0x40是黑色
        border.top_colour = 0x40 #设置框线颜色，0x40是黑色
        border.bottom_colour = 0x40 #设置框线颜色，0x40是黑色
        style1.borders = border
        # 设置单元格对齐方式
        alignment = xlwt.Alignment()
         # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.horz = 0x01
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment.vert = 0x00

        # 设置自动换行
        alignment.wrap = 1
        style1.alignment = alignment

        #********************写入EXCEL***************************************
        for i in range(0,hangshu-1):
            row=i #行位置
            col=0#列位置
            value=(str(harddisk1[i]))[0:-2]+"-"+(str(harddisk2[i]))[0:-2] +" H      "+ harddisk4[i] #以固定的格式写入
            first_col=work_sheet.col(i)
            sec_col=work_sheet.col(0)
            work_sheet.col(0).width = 16 * 256
            work_sheet.write(row, col, value, style1)  # 带样式的写入
            work_book.save(r"C:\Users\scyantao\Desktop\test1\1.xlsx")  # 保存文件