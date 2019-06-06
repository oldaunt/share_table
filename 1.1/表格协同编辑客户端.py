import tkinter as tk
import random
from tkinter import messagebox
from multiprocessing.managers import BaseManager
from threading import Thread
from time import sleep

MANAGER_DOMAIN ='192.168.43.221'   #输入服务器IP地址
MANAGER_PORT = 27131
MANAGER_AUTH_KEY = b'1234'

#随机生成颜色，颜色作为客户端的唯一标识
COLOR= '#'
for i in range(6):
    COLOR+=random.choice(['0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F'])
__author__ = 'jinke18'

INFO = {'name':'yourname','color':COLOR}


_idx = 1
_place=(1,1)  #用于记录焦点单元格位置
class Cell(tk.Entry):
    def __init__(self,parent,row, column, text):#row,column从1开始
        self.row=row
        self.column=column
        self.var = tk.StringVar()
        self.var.trace("w", self.resetWidth)
        self.text = '' if not text else text
        super().__init__(parent,textvariable=self.var,
                       highlightbackground='black',
                       highlightcolor = INFO['color'],
                       highlightthickness=1,
                       takefocus=True)
        self.grid(row=row-1, column=column-1, sticky=tk.E+tk.W)
        self.var.set(self.text)
        #print(type(self.var.get()))
        self.bind("<FocusIn>",self.onFocusIn)
        self.bind("<FocusOut>",self.onFocusOut)
    def put(self,**kwargs):
        self.grid(row=self.row-1, column=self.column-1, sticky=tk.NE+tk.SW, **kwargs)

    def read(self):
        text = self.var.get()
        if not text:
            return None
        else:
            return text
    def write(self,text):
        text = '' if not text else text
        self.var.set(text)
    def resetWidth(self,*args):
        self.text = self.var.get()
        width=0
        for s in self.text:
            if '0'<s<'9'or s == ' ':
                width+=1
            elif 'a'<s<'z' or 'A'<s<'Z':
                width+=1.25
            else:
                width+=1.67
        self.width=int(width+1)
        self['width']=self.width
    def resetAll(self,text,attr): #attr属性字典
        if attr['highlightbackground'] !=INFO['color']:
            self.write(text)
            print('write')
            if attr['highlightbackground'] != self['highlightbackground']:
                self.configure(attr)
    def onFocusIn(self,*args):
        global _place
        _place = (self.row,self.column)
        sheetdata.reset(_place,highlightbackground=INFO['color'],state='disabled')
    def onFocusOut(self,*args):
        #print(self.var.get())
        sheetdata.reset((self.row,self.column),text=self.read(),highlightbackground='black',state='normal')
        print('FocusOut,resetdata')
class sheetFrame(tk.Frame):
    def __init__(self, root):
        tk.Frame.__init__(self, root)
        self.canvas = tk.Canvas(root, borderwidth=0)
        self.frame = tk.Frame(self.canvas)
        
        self.vsb = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.vsb.pack(side="right", fill="y")
        
        self.hsb=tk.Scrollbar(root,orient='horizontal',command=self.canvas.xview)
        self.hsb.pack(side='bottom', fill='x')

        self.canvas.configure(yscrollcommand=self.vsb.set,xscrollcommand=self.hsb.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw", tags="sheetFrame")
        self.frame.bind("<Configure>", self.onFrameConfigure)
        self.cells={}
        self.canvas.bind_all('<MouseWheel>',self.onMouseWheel)
    def onMouseWheel(self,event):
        self.canvas.yview_scroll(-1*(event.delta//120),'units')
    def setcell(self,place,text):
        self.cells[place]=Cell(self.frame,*place,text)
    def onFrameConfigure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

class dataManager(BaseManager):
    pass

dataManager.register('sheetdata')

class DataClient():
    def __init__(self, domain, port, auth_key):
        self.domain = domain
        self.port = port
        self.auth_key = auth_key
        self.datamanager = dataManager(address=(self.domain, self.port), authkey=self.auth_key)
        self.datamanager.connect()
        print('connected to server')
    def get_sheetdata(self):
        self.sheetdata=self.datamanager.sheetdata()
        return self.sheetdata
def auto_reset():
    global _idx
    while 1:
        try:
            place = sheetdata.get_resetplace(_idx)
        except ConnectionResetError:
            messagebox.showerror('连接服务器失败','服务器关闭，无法同步表格更改')
            root.destroy()
        if place:
            celldata = sheetdata.get_datadict(place)
            sheetframe.cells[place].resetAll(celldata['text'],celldata['attr'])
            _idx+=1
        else:
            sleep(2)
def on_close():
    if messagebox.askokcancel('提示','确认退出吗？'):
        try:
            sheetdata.reset(_place,highlightbackground='black',state='normal')
        except ConnectionResetError:
            pass
        root.destroy()
if __name__ == '__main__':
    root=tk.Tk()
    try:
        client=DataClient( MANAGER_DOMAIN,MANAGER_PORT, MANAGER_AUTH_KEY)
    except ConnectionRefusedError:
        import traceback
        root.withdraw()
        messagebox.showerror('连接服务器失败','请启动服务器')
        traceback.print_exc()
    sheetdata=client.get_sheetdata()
    _idx = sheetdata.get_max_idx()
    print('_idx',_idx)
    root.title('表格协同编辑') 
    root.protocol('WM_DELETE_WINDOW',on_close)
    root.geometry('1200x750')
    sheetframe = sheetFrame(root)
    
    datadict,merged_cells=sheetdata.get_datadict('get_merge')
    print(len(datadict))
    for place in datadict.keys():
        sheetframe.setcell(place,datadict[place]['text'])
    for cell in merged_cells:
        print(cell)
        srow,scol,rowspan,colspan=cell
        for row in range(srow,srow+rowspan):
            for col in range(scol,scol+colspan):
                sheetframe.cells[(row,col)].grid_forget()
        sheetframe.cells[(srow,scol)].put(rowspan = rowspan,columnspan=colspan)
    reset_thread=Thread(target=auto_reset,args=())
    reset_thread.setDaemon(True)
    reset_thread.start()
    
    root.mainloop()

