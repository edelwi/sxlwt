# -*- coding: cp1251 -*-
# Simple xlwt v0.02w
# by Evgeniy Semenov 18 09 2012

# This product includes software developed by
# Roman V. Kiseliov <roman@kiseliov.ru>.

import xlwt

class sxlwt(object):
    def __init__(self, fname, sheet, encod='cp1251'):
        self.fname=fname
        self.wb=xlwt.Workbook(encoding=encod)
        self.ws=self.wb.add_sheet(sheet)
        self.row=0

    def setFontStyle(self, fontName, colourInd, Bold=False):
        font = xlwt.Font()
        font.name = fontName
        font.colour_index = colourInd
        font.bold = Bold
        style = xlwt.XFStyle()
        style.font = font
        return style
    def addSheet(self,SheetName):
        self.ws=self.wb.add_sheet(SheetName)
        self.row=0
#    def writeTitle(self, style):
#        self.ws.write(self.row, 0, 'CB num', style)
#        self.ws.write(self.row, 1, 'UPS num', style)
#        self.ws.write(self.row, 2, 'SUN num', style)
#        self.ws.write(self.row, 3, 'Company Name', style)
#        self.ws.write(self.row, 4, 'Address', style)
#        self.row+=1
    def writerow(self, style, *some):
        col=0
        for i in some:
            self.ws.write(self.row, col, i, style); col+=1
        self.row+=1
    def writerowAsL(self, style, lst):
        col=0
        for i in lst:
            self.ws.write(self.row, col, i, style); col+=1
        self.row+=1

    def setColWidth(self, *width):
        colnum=0
        for i in width:
            self.ws.col(colnum).width = i; colnum+=1
    def save(self):
        self.wb.save(self.fname)
    def write(self,row,col,data,style):
        self.ws.write(row, col, data, style)

if __name__=="__main__":
    x=sxlwt('test.xls','sheet')
    style=x.setFontStyle('Times New Roman',4,True)
    x.writerow(style,'Поле1','Lobuda','Vsyachina')
    style=x.setFontStyle('Times New Roman',0,False)
    x.writerow(style,'1')
    x.writerow(style,'','1')
    x.writerow(style,'','','1')
    x.setColWidth(3000,10000,5000)
    x.save()