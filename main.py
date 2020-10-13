#import os
import sys
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageWin #Image, ImageFont, ImageDraw, PngImagePlugin
from barcode.writer import ImageWriter
from barcode.writer import *
import barcode
import win32api, win32print, win32ui, win32con

#for pyinstaller + Label for bach scan.txt
def resource_path(relative_path):
       """ Get absolute path to resource, works for dev and for PyInstaller """
       try:
           # PyInstaller creates a temp folder and stores path in _MEIPASS
           base_path = sys._MEIPASS
       except Exception:
           base_path = os.path.abspath(".")
       return os.path.join(base_path, relative_path)

font_style= resource_path("Arial.tft")

#Chose printer
def choose(event):
    return select.get()

#chose leather type
select_type = ['Dakota', 'Saddle', 'Spalt', 'Merino', 'Perlnappa', 'Feinnappa', 'Vachette', 'NE-51']
def check_select_type(event):
    return str(leather_type.get())

#Make barcode
def Ean():
    EAN = barcode.get_barcode_class('Code128')
    ean = barcode.get('Code128', str(leather_field.get()), writer=barcode.writer.ImageWriter())
    ean.save(tempfile.gettempdir() + "\\barcode", {"module_width":0.25, "module_height":5, "font_size": 14, "text_distance": 3, "quiet_zone": 3})

#Work run
def get_bach():
    g_bach = str(leather_field.get())
    g_loss = str(loss_calc_field.get())
    g_type = str(leather_type.get())
    barcode_path = os.path.join(tempfile.gettempdir(), 'barcode.png')
    file_path = os.path.join(tempfile.gettempdir(), 'file.png')
    if os.path.isfile(barcode_path):
        os.remove(barcode_path)
    if os.path.isfile(file_path):
        os.remove(file_path)
    if g_bach =='':
        messagebox.showinfo("Партія шкіри:", "Поле партії не може бути пустим!")
    elif not g_type in select_type:
        messagebox.showinfo("Тип шкіри:", "Оберіть тип шкіри!")
    elif g_loss == '':
        messagebox.showinfo("№ Розрахунку витрат:", "Поле розрахунку витрат не може бути пустим!")
    elif g_bach != '' and g_loss != '':
        if __name__ == '__main__':
            Ean()
            subprocess.call(['attrib', '+h', barcode_path])
        g_etik=count_field.get()
        if int(g_etik) > 99:
            messagebox.showinfo("Увага!","Перевірте коректність вводу кількості етикеток!")
        count_field.delete(0, tk.END)
        count_field.insert(0,1)

        txt = Image.open(barcode_path)
        W,H = txt.size
        fnt = ImageFont.truetype("arial.ttf", 36)
        fnt_t = ImageFont.truetype("arial.ttf", 18)
        d = ImageDraw.Draw(txt)
        w,h = d.textsize(g_loss)
        w1,h1 = d.textsize(g_type)
        d.text(((float(W/2)-(w*1.5)), (float(H/3)+(h*2))), g_loss, font=fnt, fill=(0,0,0,255))
        d.text(((float(W / 2) - (w1 * 0.7)), (float(H / 2.2) + (h1 * 5.1))), g_type, (0,0,0), font=fnt_t )# fill=(255, 255, 255, 255)
        txt.save(file_path, "PNG")
        subprocess.call(['attrib', '+h', file_path])
        txt.close()
        file_name = file_path 
        i = int(g_etik)
        if int(g_etik) <= 99:
            while i > 0:
                # Add loss calculation number to barcode
                # Print finish result
                HORZRES = 8
                VERTRES = 10

                # LOGPIXELS = dots per inch
                LOGPIXELSX = 88
                LOGPIXELSY = 90

                # PHYSICALWIDTH/HEIGHT = total area
                PHYSICALWIDTH = 110
                PHYSICALHEIGHT = 111

                # PHYSICALOFFSETX/Y = left / top margin
                PHYSICALOFFSETX = 112
                PHYSICALOFFSETY = 113

                #printer_name = win32print.GetDefaultPrinter()
                #file_name = "file.png"

                # You can only write a Device-independent bitmap
                #  directly to a Windows device context; therefore
                #  we need (for ease) to use the Python Imaging
                #  Library to manipulate the image.
                #
                # Create a device context from a named printer
                #  and assess the printable size of the paper.
                hDC = win32ui.CreateDC()
                hDC.CreatePrinterDC(select.get())
                printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
                printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
                printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)

                # Open the image, rotate it if it's wider than
                #  it is high, and work out how much to multiply
                #  each pixel by to get it as big as possible on
                #  the page without distorting.
                bmp = Image.open(file_name)
                if bmp.size[0] > bmp.size[1]:
                    bmp = bmp.rotate(0)

                ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
                scale = min(ratios)

                # Start the print job, and draw the bitmap to
                #  the printer device at the scaled size.
                hDC.StartDoc(file_name)
                hDC.StartPage()

                dib = ImageWin.Dib(bmp)
                scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
                x1 = int((printer_size[0] - scaled_width) / 2)
                y1 = int((printer_size[1] - scaled_height) / 2)
                x2 = x1 + scaled_width
                y2 = y1 + scaled_height
                dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

                hDC.EndPage()
                hDC.EndDoc()
                hDC.DeleteDC()
                i -= 1

                bmp.close()
                loss_calc_field.delete(0, tk.END)
                leather_field.delete(0, tk.END)
                leather_type.delete(0, tk.END)
            os.remove(barcode_path)
            os.remove(file_path)

#-----head--------------------------------------------------------------------------------------------------------------
root = tk.Tk()
root.title('')
icon = resource_path("favicon.ico")
root.iconbitmap(icon)
root.geometry('400x600+100+100')
root.resizable(0,0)
root.configure(bg='white')
frame_tick = tk.Frame(root,bg = 'white')
frame_tick.pack()
logo = resource_path("biz-ex-logo.png")
img = tk.PhotoImage(file=logo)
main_logo = tk.Label(frame_tick, image=img)
main_logo.configure(bg='white', anchor=tk.CENTER, width=250)
main_logo.grid(ipadx=70,ipady=1,column=0,row=0)

#-----Printer select menu-----------------------------------------------------------------------------------------------
printers = win32print.EnumPrinters ( win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
res =[printer[2] for printer in printers]

frame_print_label = tk.LabelFrame(root, text="Виберіть принтер з списку:",padx=27,pady=10, bg="white",font=(font_style, 11))
frame_print_label.place(x=10,y=50)

select = ttk.Combobox(frame_print_label,values = res)
select.set(res[0])
select.grid(ipadx=90,ipady=2,column=0,row=2)
select.bind("<<ComboboxSelected>>", choose)

#-----Entry data menu-----------------------------------------------------------------------------------------------
leather_label = tk.Label(root, text="Партія шкіри:")
leather_label.configure(bg='white', width=250,anchor=tk.W, font=(font_style, 20))
leather_label.place(x=10,y=115)

leather_field = tk.Entry(root)
leather_field.configure(width=9,bd =5,font=(font_style,25))
leather_field.place(x=15,y=160)
def foo_bach(e):
    s = leather_field.get().strip()
    s = s[-1] if s in range(0,9) else ''
    leather_field.delete ('9',tk.END)
    leather_field.insert(tk.INSERT,s)
leather_field.bind('<KeyRelease>',foo_bach)

frame_label_type = tk.LabelFrame(root, text="Вкажіть тип шкіри:",padx=10,pady=5, bg="white",font=(font_style, 11))
frame_label_type.place(x=200,y=152)

leather_type = ttk.Combobox(frame_label_type,values = select_type)

leather_type.grid(padx=12,pady=2,column=1,row=15)
leather_type.bind("<<ComboboxSelected>>", check_select_type)

loss_calc_label = tk.Label(root, text="№ Розрахунку витрат:")
loss_calc_label.configure(bg='white', width=250, anchor=tk.W,font=(font_style, 20))
loss_calc_label.place(x=10,y=215)

loss_calc_field = tk.Entry(root)
loss_calc_field.configure(width=6,bd =5,font=(font_style,25))
loss_calc_field.place(x=15,y=260)
def foo_loss(e):
    s = loss_calc_field.get().strip()
    s = s[-1] if s in range(0,6) else ''
    loss_calc_field.delete ('6',tk.END)
    loss_calc_field.insert(tk.INSERT,s)
loss_calc_field.bind('<KeyRelease>',foo_loss)

count_label = tk.Label(root, text="Кількість етикеток:")
count_label.configure(bg='white', width=250, anchor=tk.W,font=(font_style, 20))
count_label.place(x=10,y=320)

count_field = tk.Spinbox(root,from_=1,to= 99,font=(font_style, 20))
count_field.configure(bd = 5,width=5)
count_field.place(x=15,y=370)

#-----Run job -----------------------------------------------------------------------------------------------
btn = tk.Button(root, text="ДРУК", command=get_bach)
btn.configure(width=20, font=(font_style, 20), bd=3, fg="#394F8C", activeforeground="red")
btn.place(x=35,y=470)

def about():
    a = tk.Toplevel()
    a.update_idletasks()
    width = 400
    height = a.winfo_height()
    x = (a.winfo_screenwidth() // 2) - (width // 2)
    y = (a.winfo_screenheight() // 2) - (height // 2)
    a.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    a.title('About:')
    a['bg'] = 'white'
    a.overrideredirect(True)
    tk.Label(a,bg='white',width=400,font=(font_style, 10), text="ІТ department Bader Ukraine\n\n Email: Yevhen.Konyukhov@bader-leather.com").pack(expand=1)
    a.after(5000, lambda: a.destroy())

btn_about = tk.Button(root, text="About",command=about)
btn_about.configure(width=10, font=(font_style, 10), bd=3, fg="#394F8C", activeforeground="red")
btn_about.place(x=150,y=550)

root.mainloop()


