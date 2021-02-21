import tkinter as tk
import PyPDF2 as pdf
from tkinter.filedialog import asksaveasfilename
from os import startfile
import re
import traceback


class Application(tk.Frame):
    def __init__(self, geometry, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("PyPDF GUI")
        self.pack()
        self.create_widgets()
        self.geometry = geometry

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
        self.space_label1 = tk.Label(self, text="",width=25)
        self.space_label1.grid(row=0,column=3)
        self.space_label2 = tk.Label(self, text="")
        self.space_label2.grid(row=2,column=3)
        self.space_label3 = tk.Label(self, text="")
        self.space_label3.grid(row=11,column=3)
        self.message_label = tk.Label(self, text="",fg="red")
        self.message_label.grid(row=9,column=1,columnspan=4)

        #Help Button
        self.help_button = tk.Button(self,text="Help",width=10,command=self.launch_help)
        self.help_button.grid(row=12, column=3)

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
        self.message_label["text"] = ""

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
        self.message_label["text"] = ""

    def launch_help(self):
        self.help_window = tk.Toplevel(self.master)
        self.help_window.geometry(self.geometry)
        self.help_window.title("PyPDF GUI Help")

        self.combine_help_header = tk.Label(self.help_window, text="Combine PDFs", anchor="w", justify="left",font=('Verdana','9','bold', 'underline'))
        self.combine_help_header.grid(row=1,column=1, sticky="NW")
        self.split_help_header = tk.Label(self.help_window, text="Split PDFs", anchor="w", justify="left",font=('Verdana','9','bold', 'underline'))
        self.split_help_header.grid(row=1,column=2, sticky="NW")
        combine_help_string =  ("Entry Fields\n"
                                "First PDF: double click the text box to search for the first file you wish to combine. \"Open\" the file and the path will populate the text box.\n\n"
                                "Second PDF: same as First PDF, but this PDF will be appended to the end of the first PDF.\n\n"
                                "New PDF: double click the text box to search for the file location you want to save the new PDF. Type a name to save the file as and click Save. This file cannot overwrite either file being combined.")
        split_help_string =  ("Entry Fields:\n"
                              "PDF: double click the text box to search for the file you wish to split. \"Open\" the file and the path will populate the text box.\n\n"
                              "Split Interval: enter where you wish the PDF to be split at. There are 3 ways to split a document.\n"
                              " - Enter a number to split the document at that page.\n"
                              " - Enter multiple numbers separated by commas to split the document at each page listed.\n"
                              " - Enter a number followed by an asterisk (*) to split the PDF at every nth page.\n"
                              "If the interval doesn't follow one of the formats, it will not process.")
        self.combine_help = tk.Label(self.help_window, text=combine_help_string, width=42, anchor="w", justify="left",wraplength=295)
        self.combine_help.grid(row=2,column=1, sticky="NW")
        self.split_help = tk.Label(self.help_window, text=split_help_string, width=42, anchor="w", justify="left",wraplength=295)
        self.split_help.grid(row=2,column=2, sticky="NW")
        
        go_help_string = "Click Go to process the combination/split. If it is successful, all text fields will be cleared for the process. If it does not complete, a red message will appear above the Go button."
        self.go_help = tk.Label(self.help_window, text=go_help_string, justify="center",wraplength="500")
        self.go_help.grid(row=3,column=1,columnspan=2)

        self.close_help = tk.Button(self.help_window,text="Close",width=10,command=self.help_window.destroy)
        self.close_help.grid(row=4,column=1,columnspan=2)



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
            self.message_label["text"] = "Cannot split at first page or Interval out of bounds"
            traceback.print_exc()
        except FileNotFoundError as error:
            self.message_label["text"] = "File name(s) missing or incorrect"
            traceback.print_exc()
        except IOError as error:
            self.message_label["text"] = "Split Interval format not valid"
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
            self.message_label["text"] = "File name(s) missing or incorrect"
            traceback.print_exc()
        except NameError as error:
            self.message_label["text"] = "New file name must be different from files to be combined"
            traceback.print_exc()


root = tk.Tk()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = 1000
y = 500
x_offset = int((screen_width-x)/2)
y_offset = int((screen_height-y)/2)
geometry_string = str(x) + 'x' + str(y) + '+' + str(x_offset) + '+' + str(y_offset)
root.geometry(geometry_string)
root.resizable(0,0)
app = Application(geometry=geometry_string,master=root)
app.mainloop()
