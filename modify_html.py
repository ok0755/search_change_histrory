#coding=gb18030
import xlrd
import string
import win32api
from Tkinter import *
import os
import arrow
import datetime

class CreateHtml(object):
    def __init__(self,model):
        self.model=model
        self.ar=[]
        self.book3=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2016\ECR編號跟進-2016.xls','ECR']
        self.book4=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2017\ECR编号跟进-2017.xls','ECR']
        self.book5=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2018\ECR编号跟进-2018.xls','ECR']
        self.book6=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2016\PCR編號跟進-2016.xls','PCR']
        self.book7=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2017\PCR編號跟進-2017.xls','PCR']
        self.book8=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\IMP_APE\APE Report File\ECR & PCR 編號跟進.XLS\2018\PCR編號跟進-2018.xls','PCR']
        self.book_lists=[self.book3,self.book4,self.book5,self.book6,self.book7,self.book8]
        self.CreateLists()

    def CreateLists(self):
        for book in self.book_lists:
            k=0
            wb=xlrd.open_workbook(book[0],formatting_info=True)
            sh=wb.sheet_by_name(book[1])
            for rows_ in sh.col_values(4):
                if self.model in rows_:
                    try:                     #判断有无链接
                        link=sh.hyperlink_map.get((k,0))
                        lists_link=link.url_or_path.encode('gb18030')
                        lists_value=sh.cell(k,0).value.encode('gb18030')
                        column_1='<tr bgcolor="#F5F5DC"><td width="60"><a href="{}">{}</a></td>'.format(lists_link,lists_value)
                    except:
                        column_1='<tr bgcolor="#FF99CC"><td width="60">'+sh.cell(k,0).value.encode('gb18030')+'</td>'
                    try:
                        date=xlrd.xldate.xldate_as_datetime(sh.cell(k,1).value, 0)   #单元格日期
                        da=arrow.get(date)
                    except:
                        da='1949/10/1'
                    column_2='<td width="80">{}</td>'.format(da.format('YYYY-M-D'))
                    column_3='<td width="50">{}</td>'.format(sh.cell(k,2).value.encode('gb18030'))
                    column_4='<td width="50">{}</td>'.format(sh.cell(k,3).value.encode('gb18030'))
                    column_5='<td width="200">{}</td>'.format(sh.cell(k,4).value.encode('gb18030'))
                    column_6='<td width="200">{}</td>'.format(sh.cell(k,5).value.encode('gb18030'))
                    column_7='<td width="40">{}</td></tr>'.format(sh.cell(k,6).value.encode('gb18030'))
                    self.ar.append([column_1,column_2,column_3,column_4,column_5,column_6,column_7])
                k+=1
        self.write_html_header(self.ar)

    def write_html_header(self,ar):
        f=open('d:\\result.html','w')
        html='''
        <html>
        <head>
        <base target="_blank"/>
        <style type="text/css">
        a:link{text-decoration:none;}
        a:hover{color:#FF00FF;}
        </style>
        </head>
        <body>
        <table width="1200">
        '''
        f.write(html)
        for k in ar:
            for kk in k:
                f.write(kk)
        f.write('</table></body></html>')
        f.close()

def cmd_exe(event=None):
    motor_model=string.upper(e.get())
    CreateHtml(motor_model)
    win32api.ShellExecute(0,'open',r'd:\result.html','','',1)
    #os.popen(r'd:\result.html')
    root.quit
    Tk.quit

if __name__=='__main__':
    root=Tk()
    root.geometry('350x30+440+400')
    root.resizable=FALSE
    root.title(u'更改记录查询')
    Label(root,text=u'输入马达型号:').grid(sticky=W,row=1,column=0)
    keyword=StringVar()
    e=Entry(root,textvariable=keyword,width=30)
    e.grid(sticky=W,row=1,column=1)
    Entry.focus_set(e)
    e.bind("<Return>",cmd_exe)
    root.mainloop()


