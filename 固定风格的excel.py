#excel 固定格式转变，把一些excel数据转化成自己喜欢风格的表格
# -*- coding: utf-8 -*- 
import xdrlib
import xlrd 
import xlwt

import uniout

import sys  
   
def open_excel(file='file.xls'): 
    try: 
        data = xlrd.open_workbook(file) 
        return data 
    except Exception,e: 
        print str(e) 

def excel_table_byname(file='file.xls',colnameindex=0,by_name=u'Sheet1',sheetNo=0): 
    data = open_excel(file) 
    if sheetNo==0:
        table =data.sheet_by_index(sheetNo)
    else:
        table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数 
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


def save_my_excel(tables,filenames="new_excel.xls",Worksheet="new_sheetname"):
    result=tables[0]
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(Worksheet)
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
    workbook.save(filenames)

def main(): 
    tables = excel_table_byname(file='query_result.xls',by_name=u"Tabl5ib Dataset",sheetNo=0) #表名和表序号填一个就可以
    save_my_excel(tables,filenames="new_excel.xls",Worksheet=u"Tablib Dataset")

if __name__=="__main__": 
    main()








