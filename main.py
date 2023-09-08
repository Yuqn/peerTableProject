import xlwt
import calendar
import time
from datetime import date,timedelta
# import patterns as patterns

# 获取时间戳
utc_time = calendar.timegm(time.gmtime())

# 创建表
workBook = xlwt.Workbook(encoding='utf-8')
workSheet = workBook.add_sheet('同行表')

# 封装首行样式
def define_style():
    font = xlwt.Font() # 字体类型
    font.colour_index = 0 # 字体颜色
    font.height = 20 * 16 # 字体大小，16为字号，20为衡量单位
    font.name='宋体' # 字体类型
    font.italic = False # 取消字体斜体
    font.bold = False # 字体加粗
    alignment = xlwt.Alignment() # 设置单元格对齐方式
    alignment.horz = 0x02 # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    # borders.right = 2
    borders.top = 1
    borders.bottom = 1
    pattern = xlwt.Pattern() # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN 
    pattern.pattern_fore_colour = 34 # 背景颜色
    
    my_style = xlwt.XFStyle()
    my_style.font = font  # 设置字体
    my_style.alignment = alignment  # 设置对齐方式
    my_style.borders = borders  # 设置边框
    my_style.pattern = pattern  # 设置背景颜色
    return my_style

# 封装2、3行样式
def define_style_two():
    font = xlwt.Font() # 字体类型
    font.colour_index = 0 # 字体颜色
    font.height = 20 * 10 # 字体大小，16为字号，20为衡量单位
    font.name='宋体' # 字体类型
    font.italic = False # 取消字体斜体
    font.bold = True # 字体加粗
    alignment = xlwt.Alignment() # 设置单元格对齐方式
    alignment.horz = 0x02 # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    pattern = xlwt.Pattern() # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN 
    pattern.pattern_fore_colour = 40 # 背景颜色
    
    my_style_two = xlwt.XFStyle()
    my_style_two.font = font  # 设置字体
    my_style_two.alignment = alignment  # 设置对齐方式
    my_style_two.borders = borders  # 设置边框
    my_style_two.pattern = pattern  # 设置背景颜色
    return my_style_two

# 封装同行商家行样式
def define_style_three():
    font = xlwt.Font() # 字体类型
    font.colour_index = 0 # 字体颜色
    font.height = 20 * 10 # 字体大小，16为字号，20为衡量单位
    font.name='宋体' # 字体类型
    font.italic = False # 取消字体斜体
    font.bold = False # 字体加粗
    alignment = xlwt.Alignment() # 设置单元格对齐方式
    alignment.horz = 0x02 # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    #pattern = xlwt.Pattern() # 设置背景颜色的模式
    #pattern.pattern = xlwt.Pattern.SOLID_PATTERN 
    #pattern.pattern_fore_colour = 40 # 背景颜色
    
    my_style_three = xlwt.XFStyle()
    my_style_three.font = font  # 设置字体
    my_style_three.alignment = alignment  # 设置对齐方式
    my_style_three.borders = borders  # 设置边框
    #my_style_three.pattern = pattern  # 设置背景颜色
    return my_style_three

# 封装预定、总数行样式
def define_style_four():
    font = xlwt.Font() # 字体类型
    font.colour_index = 0 # 字体颜色
    font.height = 20 * 10 # 字体大小，16为字号，20为衡量单位
    font.name='宋体' # 字体类型
    font.italic = False # 取消字体斜体
    font.bold = False # 字体加粗
    alignment = xlwt.Alignment() # 设置单元格对齐方式
    alignment.horz = 0x02 # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    pattern = xlwt.Pattern() # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN 
    pattern.pattern_fore_colour = 50 # 背景颜色
    
    my_style_four = xlwt.XFStyle()
    my_style_four.font = font  # 设置字体
    my_style_four.alignment = alignment  # 设置对齐方式
    my_style_four.borders = borders  # 设置边框
    my_style_four.pattern = pattern  # 设置背景颜色
    return my_style_four

# 封装入住率、空房数量行样式
def define_style_five():
    font = xlwt.Font() # 字体类型
    font.colour_index = 0 # 字体颜色
    font.height = 20 * 10 # 字体大小，16为字号，20为衡量单位
    font.name='宋体' # 字体类型
    font.italic = False # 取消字体斜体
    font.bold = False # 字体加粗
    alignment = xlwt.Alignment() # 设置单元格对齐方式
    alignment.horz = 0x02 # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    pattern = xlwt.Pattern() # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN 
    pattern.pattern_fore_colour = 51 # 背景颜色
    
    my_style_five = xlwt.XFStyle()
    my_style_five.font = font  # 设置字体
    my_style_five.alignment = alignment  # 设置对齐方式
    my_style_five.borders = borders  # 设置边框
    my_style_five.pattern = pattern  # 设置背景颜色
    return my_style_five

# 设置行高
for i in range(15):
    workSheet.row(i).set_style(xlwt.easyxf('font:height 512;'))

# 设置列宽
for i in range(13):
    workSheet.col(i+1).width = 256*13

# 合并第1行到第1行的第1列到第8列，并且添加样式
mystyle = define_style()
workSheet.write_merge(0,0,0,7,label='******店',style=mystyle)

# 第二、三行数据填充
mystyletwo = define_style_two()
day = ['6-10','6-10','6-10','6-10','6-10','6-10','6-10']
week = ['星期一','星期一','星期一','星期一','星期一','星期一','星期一']
workSheet.write(1,0,"日期",style=mystyletwo)
workSheet.write(2,0,"携程",style=mystyletwo)
for i in range(len(day)):
    workSheet.write(1,i+1,day[i],style=mystyletwo)
    workSheet.write(2,i+1,week[i],style=mystyletwo)

# 设置同行商家行样式
mystylethree = define_style_three()
peer = ['嘻哈','星驿','至尚','凌月','欢腾','香蕉管'] # 同行商家
for i in range(len(peer)):
    workSheet.write(i+3,0,peer[i],style=mystylethree)

# 订阅、总数行
mystylefour = define_style_four()
dueNum = [40,40,40,40,40,40,40] # 预定数量
countNum = [40,40,40,40,40,40,40] # 总房间数
workSheet.write(9,0,"预定数量",style=mystylefour)
workSheet.write(10,0,"房间总数",style=mystylefour)
for i in range(len(dueNum)):
    workSheet.write(9,i+1,dueNum[i],style=mystylefour)
    workSheet.write(10,i+1,countNum[i],style=mystylefour)

# 入住率、空房数量行
mystylefive = define_style_five()
rate=['100%','100%','100%','100%','100%','100%','100%'] # 入住率
remain=[0,0,0,0,0,0,0] # 剩余量
workSheet.write(11,0,"入住率",style=mystylefive)
workSheet.write(12,0,"空房数量",style=mystylefive)
for i in range(len(rate)):
    workSheet.write(11,i+1,rate[i],style=mystylefive)
    workSheet.write(12,i+1,remain[i],style=mystylefive)

# 填充末尾行
nowToday = date.today() # 获取时间
maker = '余' # 制表人
process = '*' # 审核人
workSheet.write(13,0,"制表时间",style=mystyletwo)
workSheet.write_merge(13,13,1,3,label=str(nowToday),style=mystyletwo)
workSheet.write(13,4,"制表人",style=mystyletwo)
workSheet.write(13,5,maker,style=mystyletwo)
workSheet.write(13,6,"审核人",style=mystyletwo)
workSheet.write(13,7,process,style=mystyletwo)

# 保存
savePath = './' + str(utc_time) + 'peer.xls'
workBook.save(savePath)

























'''
import xlwt
def define_style():
    font = xlwt.Font()# 字体类型
    font.name = 'name Times New Roman'# 字体颜色
    font.colour_index = 1# 字体大小，16为字号，20为衡量单位
    font.height = 20 * 16# 字体加粗
    font.bold = False# 下划线
    font.underline = True# 斜体字
    font.italic = True# 设置单元格对齐方式
    alignment = xlwt.Alignment()# 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02# 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01# 设置自动换行
    alignment.wrap = 1# 设置边框
    borders = xlwt.Borders()# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 2
    borders.top = 3
    borders.bottom = 4
    borders.left_colour = 1
    borders.right_colour = 2
    borders.top_colour = 3
    borders.bottom_colour = 4# 设置列宽，一个中文等于两个英文等于两个字符，11为字符数，256为衡量单位
    sheet.col(1).width = 11 * 256# 设置背景颜色
    pattern = xlwt.Pattern()# 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN# 背景颜色
    pattern.pattern_fore_colour = 3# 初始化样式
    my_style = xlwt.XFStyle()
    my_style.font = font  # 设置字体
    my_style.alignment = alignment  # 设置对齐方式
    my_style.borders = borders  # 设置边框
    my_style.pattern = pattern  # 设置背景颜色
    return my_style
if__name__ == '__main__':
book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
mystyle = define_style()
sheet.write(0, 0, u'(0,0)', mystyle) # 横坐标，纵坐标，内容，样式book.save('my_excel.xlsx')



'''


















