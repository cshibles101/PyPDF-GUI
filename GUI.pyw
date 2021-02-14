import tkinter as tk
import PyPDF2 as pdf
from tkinter.filedialog import asksaveasfilename
from os import startfile
import re
# import traceback


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("PyPDF GUI")
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        # self.grid_columnconfigure(5)
        # self.grid_rowconfigure(12)

        #Combine/Split Radio Buttons
        self.x = tk.StringVar(value="combine")
        self.combine = tk.Radiobutton(self, text="Combine", variable=self.x, value="combine", width =15, command=self.on_radio_change)
        self.split = tk.Radiobutton(self, text="Split", variable=self.x, value="split", width=15, command=self.on_radio_change)
        self.combine.grid(row=1,column=1)
        self.split.grid(row=1,column=4)

        #Combine PDF Fields
        self.first_pdf = tk.Label(self, text="First PDF")
        self.first_pdf.grid(row=3,column=1,sticky="W")
        self.first_pdf_entry = tk.Entry(self)
        self.first_pdf_entry.grid(row=4,column=1)
        self.first_pdf_entry.bind("<Double-Button-1>", self.on_entry_click)

        self.second_pdf = tk.Label(self, text="Second PDF")
        self.second_pdf.grid(row=5,column=1,sticky="W")
        self.second_pdf_entry = tk.Entry(self)
        self.second_pdf_entry.grid(row=6,column=1)
        self.second_pdf_entry.bind("<Double-Button-1>", self.on_entry_click)

        self.target_pdf = tk.Label(self, text="New PDF")
        self.target_pdf.grid(row=7,column=1,sticky="W")
        self.target_pdf_entry = tk.Entry(self)
        self.target_pdf_entry.grid(row=8,column=1)
        self.target_pdf_entry.bind("<Double-Button-1>", self.on_entry_click)

        #Split PDF Fields
        self.split_pdf = tk.Label(self, text="PDF")
        self.split_pdf.grid(row=3,column=4,sticky="W")
        self.split_pdf_entry = tk.Entry(self)
        self.split_pdf_entry.grid(row=4,column=4)
        self.split_pdf_entry.bind("<Double-Button-1>", self.on_entry_click)

        self.interval = tk.Label(self, text="Split Interval")
        self.interval.grid(row=5,column=4,sticky="W")
        self.interval_entry = tk.Entry(self)
        self.interval_entry.grid(row=6,column=4)

        #Go Button
        self.go_button = tk.Button(self,text="GO",width=15,command=self.go_button_press)
        self.go_button.grid(row=10,column=3)

        #Labels for grid spacing purposes
        self.spacelabel1 = tk.Label(self, text="",width=25)
        self.spacelabel1.grid(row=0,column=3)
        self.spacelabel2 = tk.Label(self, text="")
        self.spacelabel2.grid(row=2,column=3)
        self.messagelabel = tk.Label(self, text="",fg="red")
        self.messagelabel.grid(row=9,column=1,columnspan=4)


        #Calling method to disable based on default radiobutton chosen
        self.disable_entries(self.x.get())
    
    def go_button_press(self):
        if self.x.get() == "combine":
            self.combine_pdf_method()
        else:
            self.split_pdf_method()

    #Method used to open Windows File Explorer when a user clicks in the entry fields
    def on_entry_click(self,event):
        data = [("PDF","*.pdf")]
        filename = asksaveasfilename(filetypes = data, defaultextension = data, confirmoverwrite=False)
        root.update()
        # print(filename)
        event.widget.delete(0,len(event.widget.get()))
        event.widget.insert(0,filename)

    def on_radio_change(self):
        self.disable_entries(self.x.get())
        self.messagelabel["text"] = ""

    def disable_entries(self, option):
        if option == "combine":
            self.split_pdf_entry["state"]="disabled"
            self.interval_entry["state"]="disabled"
            self.first_pdf_entry["state"]="normal"
            self.second_pdf_entry["state"]="normal"
            self.target_pdf_entry["state"]="normal"
        else:
            self.first_pdf_entry["state"]="disabled"
            self.second_pdf_entry["state"]="disabled"
            self.target_pdf_entry["state"]="disabled"
            self.split_pdf_entry["state"]="normal"
            self.interval_entry["state"]="normal"

    def clear_entry_fields(self):
        self.split_pdf_entry.delete(0,len(self.split_pdf_entry.get()))
        self.first_pdf_entry.delete(0,len(self.first_pdf_entry.get()))
        self.second_pdf_entry.delete(0,len(self.second_pdf_entry.get()))
        self.target_pdf_entry.delete(0,len(self.target_pdf_entry.get()))
        self.interval_entry.delete(0,len(self.interval_entry.get()))
        self.messagelabel["text"] = ""

    #PDF manipulation Methods
    def split_pdf_method(self):
        filename = self.split_pdf_entry.get()
        interval_value = self.interval_entry.get()
        try:
            if re.search("(^\d+\*$)|((^(\d,)+\d$)|^\d+$)",interval_value) is None:
                raise IOError
            else:
                if ',' in interval_value: #splits at multiple specific page numbers
                    interval_values = interval_value.split(',')
                    with open(filename, 'rb') as infile:
                        reader = pdf.PdfFileReader(infile)
                        num_pages = reader.getNumPages()
                        interval_values = list(map(lambda x: int(x), interval_values))
                        interval_values.sort()
                        list(set(interval_values))
                        for value in interval_values:
                            if value <= 1 or value > num_pages:
                                raise IndexError
                        writer = pdf.PdfFileWriter()
                        writer.addPage(reader.getPage(0))
                        for page in range(1,num_pages):
                            if page+1 in interval_values:
                                if interval_values.index(page+1)-1 == -1:
                                    number = 1
                                else:
                                    number = interval_values[interval_values.index(page+1)-1]
                                destination = filename[:-4] + "_" + str(number) + ".pdf"
                                with open(destination, 'wb') as outfile:
                                    writer.write(outfile)
                                writer = pdf.PdfFileWriter()
                            writer.addPage(reader.getPage(page))
                        destination = filename[:-4] + "_" + str(interval_values.pop()) + ".pdf"
                        with open(destination, 'wb') as outfile:
                            writer.write(outfile)

                elif '*' in interval_value: #includes repeating splits at a set interval
                    interval_value = int(interval_value[:-1])
                    with open(filename, 'rb') as infile:
                        reader = pdf.PdfFileReader(infile)
                        num_pages = reader.getNumPages()
                        if interval_value == 0 or interval_value > num_pages:
                            raise IndexError
                        writer = pdf.PdfFileWriter()
                        writer.addPage(reader.getPage(0))
                        for page in range(1,num_pages):
                            if page%interval_value == 0:
                                destination = filename[:-4] + "_" + str(page+1-interval_value) + ".pdf"
                                with open(destination, 'wb') as outfile:
                                    writer.write(outfile)
                                writer = pdf.PdfFileWriter()
                            writer.addPage(reader.getPage(page))
                        if num_pages%interval_value == 0:
                            number = num_pages - interval_value + 1
                        else:
                            number = num_pages - ((num_pages % interval_value) - 1)
                        destination = filename[:-4] + "_" + str(number) + ".pdf"
                        with open(destination, 'wb') as outfile:
                            writer.write(outfile)

                else: #just a number, one single split
                    interval_value = int(interval_value)-1
                    with open(filename, 'rb') as infile:
                        reader = pdf.PdfFileReader(infile)
                        if interval_value == 0 or interval_value > reader.getNumPages()-1:
                            raise IndexError
                        writer = pdf.PdfFileWriter()
                        for page in range(interval_value):
                            writer.addPage(reader.getPage(page))
                        destination = filename[:-4] + "_1.pdf"
                        with open(destination, 'wb') as outfile:
                            writer.write(outfile)
                        writer = pdf.PdfFileWriter()
                        for page in range(interval_value,reader.getNumPages()):
                            writer.addPage(reader.getPage(page))
                        destination = filename[:-4] + "_" + str(interval_value+1) + ".pdf"
                        with open(destination, 'wb') as outfile:
                            writer.write(outfile)

                self.clear_entry_fields()

            # (^\d+\*$)|((^(\d,)+\d$)|^\d+$)

        except IndexError as error:
            self.messagelabel["text"] = "Cannot split at first page or Interval out of bounds"
            traceback.print_exc()
        except FileNotFoundError as error:
            self.messagelabel["text"] = "File name(s) missing or incorrect"
            traceback.print_exc()
        except IOError as error:
            self.messagelabel["text"] = "Split Interval format not valid"
            traceback.print_exc()
    

    def combine_pdf_method(self):
        try:
            if self.target_pdf_entry.get() == "":
                raise FileNotFoundError
            if self.first_pdf_entry.get() == self.target_pdf_entry.get() or self.target_pdf_entry.get() == self.second_pdf_entry.get():
                raise NameError
            merger = pdf.PdfFileMerger()
            merger.append(self.first_pdf_entry.get())
            merger.append(self.second_pdf_entry.get())
            merger.write(self.target_pdf_entry.get())
            startfile(self.target_pdf_entry.get())
            self.clear_entry_fields()
        except FileNotFoundError as error:
            self.messagelabel["text"] = "File name(s) missing or incorrect"
            traceback.print_exc()
        except NameError as error:
            self.messagelabel["text"] = "New file name must be different from files to be combined"
            traceback.print_exc()


root = tk.Tk()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = 600
y = 300
x_offset = int((screen_width-x)/2)
y_offset = int((screen_height-y)/2)
geometry_string = str(x) + 'x' + str(y) + '+' + str(x_offset) + '+' + str(y_offset)
root.geometry(geometry_string)
root.resizable(0,0)
app = Application(master=root)
app.mainloop()
