# -*- coding: utf-8 -*-         
import os  
import uniout
import xlrd
import xlwt
#获取当前目录下xls文件名称
def file_name(file_dir=os.getcwd(),string="xls"):
    files_obj=[]
    for files in  os.listdir(file_dir):
        if os.path.splitext(files)[1]=="."+string:
            files_obj.append(file_dir+"/"+files)
    return(files_obj)
def open_excel(file='file.xls'): 
    try: 
        data = xlrd.open_workbook(file) 
        return data 
    except Exception,e: 
        print str(e)
#获取Excel分表sheet名称
def read_excel(file_name):
    data = xlrd.open_workbook(file_name)
    return(data.sheet_names())
#读取数据
def excel_table_byname(file='file.xls',colnameindex=0,by_name=u'Sheet1',sheetNo=8000): 
    data = open_excel(file) 
    if sheetNo<>8000:
        table =data.sheet_by_index(sheetNo)
    else:
        table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数 
    if nrows==0:
        print("该表为空")
        return("NULL")
    else:
        colnames = table.row_values(colnameindex) #某一行数据 
        for i in range(len(colnames)):
            if colnames[i]=="":
                colnames[i]="x"+str(i)
    #print colnames        
        list = [] 
        for rownum in range(1,nrows): 
            row = table.row_values(rownum) 
            if row: 
                app = {} 
                for i in range(len(colnames)): 
                    app[colnames[i]] = row[i] 
                list.append(app) 
    return list 
#保存数据
def save_my_excel(tables):
    result=tables[0]
    ####################################################################3
    font = xlwt.Font() # Create the Font
    font.name = u'微软雅黑'
    font.bold = False
    font.underline = False
    font.italic = True
#style = xlwt.XFStyle() # Create the Style
    pattern = xlwt.Pattern() # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    borders = xlwt.Borders() # Create Borders
    borders.left = xlwt.Borders.DASHED # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.right = xlwt.Borders.DASHED
    borders.top = xlwt.Borders.DASHED
    borders.bottom = xlwt.Borders.DASHED
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    style = xlwt.XFStyle() # Create the Pattern
    style.pattern = pattern # Add Pattern to Style
    style.font = font # Apply the Font to the Style
    style.borders = borders # Add Borders to Style

    font = xlwt.Font() # Create the Font
    font.name = u'微软雅黑'
    font.bold = False
    font.underline = False
    font.italic = False

    style1 = xlwt.XFStyle() # Create the Pattern
    style1.font = font # Apply the Font to the Style
    style1.borders = borders # Add Borders to Style
##############################################################################

    print "字段名称："
    i=1
    for friend in tables[0:]:
        j=1
        for name in result.keys():
        
 
            if i==1:    
                print "\b","\r",name
                worksheet.write(i, j, label = name,style=style)
            else:
                worksheet.write(i, j, label = friend[name],style=style1)
            j=j+1
        i=i+1   
   # workbook.save(filenames)
def main():
    #定义变量
    string="xls"
    y=file_name()#读取当前文件夹下xls文件
    
    for x in y:
        print "开始读取\""+x+"\"的数据"
        sheet=read_excel(x)
        print "sheet：" ,sheet
        # 读取数据
        sht=sheet[0]
        global workbook
        global worksheet
        workbook = xlwt.Workbook(encoding = 'ascii')
        
        print(x)
        for sht in sheet:
            
            worksheet= workbook.add_sheet(sht)
            tables = excel_table_byname(file=x,by_name=sht,sheetNo=8000) #表名和表序号填一个就可以
            if len(tables)<>0 and tables<>"NULL" :
                save_my_excel(tables)
        workbook.save(os.path.splitext(x)[0]+'_new.'+string)#保存
        print "更新为："+os.path.splitext(x)[0]+'_new.'+string

 


if __name__=="__main__": 
    main()
