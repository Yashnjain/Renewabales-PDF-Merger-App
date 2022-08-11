import tkinter
import customtkinter
from tkinter import messagebox,Tk
import sys,os
from PyPDF2 import PdfFileMerger
import traceback


customtkinter.set_appearance_mode("light")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

def resource_path(relative_path):
    try:
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception as e:
                raise e

def report_callback_exception(self,exc, val, tb):
        msg = traceback.format_exc()
        messagebox.showerror("Error", message=msg)
        app.update()

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.title("Biourja Renewables")
app["bg"]= "#e2e1ef"
biourjaLogo = resource_path('biourjaLogo.png')
photo = tkinter.PhotoImage(file = biourjaLogo)
app.iconphoto(False, photo)
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
width2 = 420
height2 = 190
x2 = (screen_width/2) - (width2/2)
y2 = (screen_height/2) - (height2/2)
app.geometry('%dx%d+%d+%d' % (width2, height2, x2, y2))


def on_closing():
        try:
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                app.destroy()
                sys.exit()
        except Exception as e:
            raise e
def button_function():
    try:
        button_text.set("PROCESSING")
        button.configure(state='disable')
        app.update()
        # input_invoices=pdf_merger()
        input_bol=r"J:\BioUrja Renewables\BOL's\Combined files\BOLs"
        input_invoices=r"J:\BioUrja Renewables\BOL's\Combined files\Invoices"
        output_folder=r"J:\BioUrja Renewables\BOL's\Combined files\Final pdf"
        if not os.path.exists(input_invoices):
            messagebox.showerror("ERROR",f"NO FOLDER FOUND : Location for the folder --> {input_invoices}")
            sys.exit()
        if not os.path.exists(output_folder):
            messagebox.showerror("ERROR",f"NO FOLDER FOUND : Location for the folder --> {output_folder}")
            sys.exit() 
        files = os.listdir(input_invoices)
        if len(files) == 0:
            messagebox.showerror("ERROR",f"FOLDER EMPTY : Please park files in the folder --> {input_invoices}")
            sys.exit()
        list_of_invoices = [a for a in os.listdir(input_invoices) if a.endswith(".pdf") and "_" in a]
        # list_of_bol = [a for a in os.listdir(input_bol) if a.endswith(".pdf")]
        for pdf in list_of_invoices:
            merger = PdfFileMerger()
            invoice_name=pdf
            if os.path.exists(input_invoices+"\\"+invoice_name):
                merger.append(input_invoices+"\\"+invoice_name,import_bookmarks=False)
            bol_name=pdf.split("_")[1]
            if os.path.exists(input_invoices+"\\"+bol_name):
                merger.append(input_invoices+"\\"+bol_name,import_bookmarks=False)
            # output_file_name=pdf.split(".pdf")[0]+ " " + "Merged file" + ".pdf"
            output_file_name=invoice_name
            merger.write(output_folder+"\\"+output_file_name)
            merger.close()
            if os.path.exists(input_invoices+"\\"+invoice_name):
                os.remove(input_invoices+"\\"+invoice_name)
            if os.path.exists(input_invoices+"\\"+bol_name):
                os.remove(input_invoices+"\\"+bol_name)
        files = os.listdir(input_invoices)
        if len(files) > 0:
            messagebox.showwarning("WARNING",f"There are some additional files in the folder --> {input_invoices}")
        button_text.set("Generate Merged PDF")
        button.configure(state='normal')
        messagebox.showinfo("INFO",f"PDF's Merged Sucessfully")
    except Exception as e:
        raise e

settings_frame = customtkinter.CTkFrame(app, width=50)
settings_frame.pack(fill=tkinter.X, side=tkinter.TOP, padx=2, pady=2)
settings_frame.grid_columnconfigure(0, weight=1)
settings_frame.grid_rowconfigure(3, weight=1)    

button_text=tkinter.StringVar()
button = customtkinter.CTkButton(master=app, textvariable=button_text, command=button_function,width=160,height=36)#,text_font=("SF Display",-13))
button_text.set("Generate Merged PDF")
button.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
app.protocol("WM_DELETE_WINDOW", on_closing)
Tk.report_callback_exception = report_callback_exception   
app.mainloop()