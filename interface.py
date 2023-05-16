import os
import gc
import io
import fitz
import numpy as np
import pickle as pkl
from tkinter import *
from PIL import Image
from PIL import ImageTk
from tkinter import ttk
from functools import partial
from tkinter import filedialog
from tkinter import messagebox
from win32api import GetSystemMetrics
from PyPDF2 import PdfFileMerger, PdfFileWriter, PdfFileReader

from default_vars import *
from texts import texts, language_conversion

w = GetSystemMetrics (0)
h = GetSystemMetrics (1)

class Interface(Frame):
    
    """
    Main window
    All the widgets are stored as attributes of this window
    """
    
    def __init__(self, window, **kwargs):
        Frame.__init__(self, window, width=window_width, height=window_height, bg=window_background_color, **kwargs)
        self.pack()
        self.widgets = {}
        self.params_file = 'params.txt'
        self.parent = window
        
        self.load_params()
        self.form_start_page()
    
    
    def load_params(self):
        
        self.dir = os.getcwd()
        self.params_file = os.path.join(self.dir, self.params_file)
        
        if not os.path.exists(self.params_file):
            self.params = {}
            self.params['language'] = 'French'
            self.params['default_root_dir'] = '/'
            self.params['up_arrow_img_path'] = 'up_arrow.png'
            
            with open(self.params_file, 'wb') as f:
                pkl.dump(self.params, f)
        
        else:
            with open(self.params_file, 'rb') as f:
                self.params = pkl.load(f)
    
    
    def erase_page(self):
        
        # Destroy previous widgets
        for key in list(self.widgets.keys()):
            self.widgets[key].destroy()
            del self.widgets[key]
        gc.collect()
        
        if hasattr(self, 'pages'):
            self.pages = []
 
 
    def open_file(self, multiple, filetypes):
        
        if multiple:
            filenames = filedialog.askopenfilenames(
                initialdir = self.params['default_root_dir'],
                title = texts[self.params['language']]['select_file'],
                filetypes = filetypes,
            )
            return list(filenames)
        filenames = filedialog.askopenfilename(
            initialdir = self.params['default_root_dir'],
            title = texts[self.params['language']]['select_file'],
            filetypes = filetypes,
        )
        return filenames

    
    def save_file(self):
        
        filename = filedialog.asksaveasfilename(
            initialdir = self.params['default_root_dir'],
            title = texts[self.params['language']]['save_file'],
            filetypes = [('pdf', '*.pdf')],
            defaultextension='pdf'
        )
        
        if filename[ -4 : ] != '.pdf':
            filename += '.pdf'
        
        if filename == '.pdf':
            return None
        
        return filename
    
    
    def color_list(self):
        
        for k in range(self.pdf_list.size()):
            if k % 2 == 0:
                self.pdf_list.itemconfigure(k, background=list_dark_color)
            else:
                self.pdf_list.itemconfigure(k, background=list_pale_color)
    
    
    def load_files(self, multiple=True, filetypes=[('PDF', '*.pdf')]): #, ('Images', '*.jpg *.jpeg *.JPG *.JPEG *.png'), ('PDF and Images', '*.pdf *.jpg *.jpeg *.JPG *.JPEG *.png')]):
        
        filepaths = self.open_file(multiple, filetypes)
        if filepaths != '' and filepaths != []:
            for filepath in filepaths:
                self.pdf_paths.append(filepath)
                filename = os.path.split(filepath)[-1]
                self.pdf_list.insert(self.pdf_list.size(), filename)
            
            self.color_list()
            
            folder = os.path.split(filepaths[0])[0]
            self.params['default_root_dir'] = folder
            with open(self.params_file, 'wb') as f:
                pkl.dump(self.params, f)
    
    
    def display_page(self, page_nb, zoom=1):
        
        assert page_nb > 0 and page_nb <= self.nb_pages
        
        wh, ww = 10, 20
        if self.load_pdf:
            img = self.pages[page_nb - 1]
            h, w = self.sizes[page_nb - 1]
            ratios = 31 * wh / h, 14 * ww / w
            ratio = zoom * min(ratios)
            img = img.resize((int(ratio * w), int(ratio * h)), Image.ANTIALIAS)
            img_tk = ImageTk.PhotoImage(img)
        
        self.widgets['pdf_text'].destroy()
        self.widgets['scrollbarv'].destroy()
        self.widgets['scrollbarh'].destroy()
        del self.widgets['pdf_text']
        del self.widgets['scrollbarv']
        del self.widgets['scrollbarh']
        
        self.pdf_text = Text(
            self,
            width=ww,
            height=wh,
            font=fonts['title'],
            fg=font_color,
        )
        if self.load_pdf:
            self.pdf_text.image_create(END, image=img_tk)
            self.pdf_text.image = img_tk
        self.scrollbarv = ttk.Scrollbar(self, orient=VERTICAL, command=self.pdf_text.yview)
        self.scrollbarh = ttk.Scrollbar(self, orient=HORIZONTAL, command=self.pdf_text.xview)
        self.pdf_text.configure(yscrollcommand=self.scrollbarv.set, xscrollcommand=self.scrollbarh.set)
        self.pdf_text.grid(row=3, column=0, rowspan=6, columnspan=7, sticky=(N, W, E, S))
        self.scrollbarv.grid(row=3, column=6, rowspan=6, sticky=(N, E, S))
        self.scrollbarh.grid(row=8, column=0, columnspan=7, sticky=(W, E, S))
        
        self.widgets['pdf_text'] = self.pdf_text
        self.widgets['scrollbarv'] = self.scrollbarv
        self.widgets['scrollbarh'] = self.scrollbarh
    
    
    def change_zoom(self, zoom):
        
        if hasattr(self, 'pages') and len(self.pages) > 0:
            self.display_page(self.current_page, zoom=int(zoom) / 100)
    
    
    def load_file(self, multiple=False, filetypes=[('PDF', '*.pdf')]):
        
        filepath = self.open_file(multiple, filetypes)
        self.load_pdf = True
        if filepath != '':
            self.filepath=filepath
            pdf_size = os.stat(filepath).st_size
            if pdf_size > 2e8:
                answer = messagebox.askyesno(texts[self.params['language']]['page3_size_yesno_title'], texts[self.params['language']]['page3_size_yesno_message'])
                if not answer:
                    self.load_pdf = False
                    self.nb_pages = PdfFileReader(filepath).getNumPages()
                    self.total_page_nb_sv.set('/ ' + str(self.nb_pages))
                    self.pages = []
                    self.current_page = 1
                    self.page_nb_sv.set(self.current_page)
                    self.display_page(self.current_page, zoom=int(self.zoom_scale.get()) / 100)
                    return
            
            open_pdf = fitz.open(filepath)
            self.pages, self.sizes = [], []
            for page in open_pdf:
                pix = page.get_pixmap()
                pix1 = fitz.Pixmap(pix, 0) if pix.alpha else pix
                height, width = pix1.height, pix1.width
                self.sizes.append((height, width))
                img = Image.open(io.BytesIO(pix1.tobytes("ppm")))
                self.pages.append(img)
            
            self.current_page = 1
            self.nb_pages = len(self.pages)
            self.page_nb_sv.set(self.current_page)
            self.total_page_nb_sv.set('/ ' + str(self.nb_pages))
            self.display_page(self.current_page, zoom=int(self.zoom_scale.get()) / 100)
    
    
    def file_select(self, event):
        
        selections = self.pdf_list.curselection()
        if len(selections) > 0:
            self.selection = selections[-1]
            self.pdf_list.selection_set(self.selection)
            self.pdf_list.see(self.selection)
    
    
    def merge_files(self):
        
        if len(self.pdf_paths) > 0:
            merger = PdfFileMerger(strict=False)
            for path in self.pdf_paths:
                merger.append(path)
            
            filename = self.save_file()
            if filename is not None:
                merger.write(filename)
            merger.close()
            
            messagebox.showinfo(
                title=texts[self.params['language']]['merge_complete_title'],
                message=texts[self.params['language']]['merge_complete_message'],
            )
    
    
    def split_pdf(self):
        
        if self.filepath is not None:
            
            if self.start_page is None or self.start_page < 0:
                messagebox.showerror(texts[self.params['language']]['missing_start_page_title'], texts[self.params['language']]['missing_start_page_message'])
                return
            
            if self.stop_page is None or self.stop_page < 0 or self.stop_page < self.start_page:
                messagebox.showerror(texts[self.params['language']]['missing_stop_page_title'], texts[self.params['language']]['missing_stop_page_message'])
                return
            
            if self.start_page > self.nb_pages or self.stop_page > self.nb_pages:
                messagebox.showerror(
                    texts[self.params['language']]['start_stop_too_big_title'],
                    texts[self.params['language']]['start_stop_too_big_message'][0] + str(self.nb_pages) + texts[self.params['language']]['start_stop_too_big_message'][1],
                )
                return
            
            input = PdfFileReader(open(self.filepath, "rb"), strict=False)
            output = PdfFileWriter()
            for i in range(self.start_page - 1, self.stop_page):
                output.addPage(input.getPage(i))
            
            filename = self.save_file()
            if filename is not None:
                with open(filename, "wb") as outputStream:
                    output.write(outputStream)
                
                del input
                del output
                gc.collect()

                messagebox.showinfo(
                    title=texts[self.params['language']]['split_complete_title'],
                    message=texts[self.params['language']]['split_complete_message'],
                )
            
            else:
                del output
                gc.collect()
    
    
    def move_up(self):
        
        selections = self.pdf_list.curselection()
        
        if len(selections) > 0:
            
            i = selections[0]
            
            if i > 0:
                
                self.pdf_paths[i], self.pdf_paths[i - 1] = self.pdf_paths[i - 1], self.pdf_paths[i]
                text = self.pdf_list.get(i)
                self.pdf_list.delete(i)
                self.pdf_list.insert(i - 1, text)
                self.pdf_list.selection_set(i - 1)
                self.pdf_list.see(i - 1)
                self.pdf_list.activate(i - 1)
                self.selection = i - 1
                self.color_list()
    
    
    def move_down(self):
        
        selections = self.pdf_list.curselection()
        
        if len(selections) > 0:
            
            i = selections[0]
            
            if i < self.pdf_list.size() - 1:
                
                self.pdf_paths[i], self.pdf_paths[i + 1] = self.pdf_paths[i + 1], self.pdf_paths[i]
                text = self.pdf_list.get(i)
                self.pdf_list.delete(i)
                self.pdf_list.insert(i + 1, text)
                self.pdf_list.selection_set(i + 1)
                self.pdf_list.see(i + 1)
                self.pdf_list.activate(i + 1)
                self.selection = i + 1
                self.color_list()
    
    
    def clear_pdfs(self):
        
        self.pdf_paths = []
        for i in range(self.pdf_list.size(), -1, -1):
            self.pdf_list.delete(i)
    
    
    def view_pdf(self, event):
        
        pass
    
    
    def retrieve_page_number(self, page):
        
        if page == "start":
            try:
                self.start_page = int(self.start_page_entry.get())
            except:
                pass
        
        elif page == "stop":
            try:
                self.stop_page = int(self.stop_page_entry.get())
            except:
                pass
        
        elif page == "display":
            try:
                current_page = int(self.page_nb_entry.get())
                if hasattr(self, 'pages') and self.nb_pages > 0:
                    if current_page > 0 and current_page <= self.nb_pages:
                        self.page_nb_sv.set(current_page)
                        self.current_page = current_page
                        self.display_page(self.current_page, zoom=int(self.zoom_scale.get()) / 100)
                    else:
                        self.page_nb_sv.set(self.current_page)
            except:
                pass
    
    
    def increase_page_nb(self):
        
        if hasattr(self, 'pages') and self.nb_pages > 0:
            new_current_page = self.current_page + 1
            if new_current_page <= self.nb_pages:
                self.page_nb_sv.set(new_current_page)
                self.current_page = new_current_page
    
    
    def decrease_page_nb(self):
        
        if hasattr(self, 'pages') and self.nb_pages > 0:
            new_current_page = self.current_page - 1
            if new_current_page > 0:
                self.page_nb_sv.set(new_current_page)
                self.current_page = new_current_page
    
    
    def create_language_menu(self, row, column, width):
        
        self.language_sv = StringVar(self, texts[self.params['language']]['self_language'])
        self.language_menu = ttk.Combobox(
            self,
            values=texts[self.params['language']]['language'],
            textvariable=self.language_sv,
            justify='center',
            font=fonts['normal'],
            width=width,
        )
        self.language_menu.bind('<<ComboboxSelected>>', self.choose_language)
        self.language_menu.grid(row=row, column=column)
        self.widgets['language_menu'] = self.language_menu
    
    
    def choose_language(self, event):
        
        text = self.language_menu.get()
        if text in texts[self.params['language']]['language']:
            
            self.params['language'] = language_conversion[text]
            with open(self.params_file, 'wb') as f:
                pkl.dump(self.params, f)
            self.form_start_page()
    
    def form_start_page(self):       
         
        self.erase_page()
        self.parent.title(texts[self.params['language']]['page1_suptitle'])
        
        # Widgets creation
        self.welcome_message = Label(
            self,
            text=texts[self.params['language']]['page1_welcome'],
            font=fonts["title"],
            fg=font_color,
            bg=window_background_color,
        )
        self.welcome_message.grid(row=1)
        self.widgets['welcome_message'] = self.welcome_message
        
        self.instruction_message = Label(
            self,
            text=texts[self.params['language']]['page1_instruction'],
            font=fonts["big bold"],
            fg=font_color,
            bg=window_background_color,
        )
        self.instruction_message.grid(row=3)
        self.widgets['instruction_message'] = self.instruction_message
        
        self.language_select = Label(
            self,
            text=texts[self.params['language']]['language_indication'],
            bg=window_background_color,
            font=fonts["bold"],
            fg=font_color,
            justify='center',
        )
        self.language_select.grid(row=9)
        self.widgets['language_select'] = self.language_select
        
        max_length = np.max([len(texts[self.params['language']][key]) for key in texts[self.params['language']].keys() if 'page1' in key])
        self.l0 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l1 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l2 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l3 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l4 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["normal"])
        self.l5 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        
        self.l0.grid(row=0)
        self.l1.grid(row=2)
        self.l3.grid(row=4)
        self.l4.grid(row=6)
        self.l5.grid(row=8)
        self.widgets['l0'] = self.l0
        self.widgets['l1'] = self.l1
        self.widgets['l2'] = self.l2
        self.widgets['l3'] = self.l3
        self.widgets['l4'] = self.l4
        self.widgets['l5'] = self.l5
        
        self.merge_button = Button(
            self,
            text=texts[self.params['language']]['page1_merge_button'],
            font=fonts["big bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            width=40,
            command=self.load_merge_page,
            cursor='hand2',
        )
        self.merge_button.grid(row=5)
        self.widgets['merge_button'] = self.merge_button
        
        self.split_button = Button(
            self,
            text=texts[self.params['language']]['page1_split_button'],
            font=fonts["big bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            width=40,
            command=self.load_split_page,
            cursor='hand2',
        )
        self.split_button.grid(row=7)
        self.widgets['split_button'] = self.split_button
        
        self.create_language_menu(10, 0, 30)
        
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)
        self.rowconfigure(4, weight=1)
        self.rowconfigure(5, weight=1)
        self.rowconfigure(6, weight=1)
        self.rowconfigure(7, weight=1)
        self.rowconfigure(8, weight=1)
        self.rowconfigure(9, weight=1)
        self.rowconfigure(10, weight=1)
    
    
    def load_merge_page(self):
        
        self.erase_page()
        self.parent.title(texts[self.params['language']]['page2_suptitle'])
        
        self.pdf_paths = []
        
        self.title = Label(
            self,
            text=texts[self.params['language']]['page2_title'],
            font=fonts["title"],
            fg=font_color,
            bg=window_background_color,
        )
        self.title.grid(row=1, columnspan=5)
        self.widgets['title'] = self.title
        
        max_length = np.max([len(texts[self.params['language']][key]) for key in texts[self.params['language']].keys() if 'page2' in key])
        self.l0 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l1 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l2 = Label(self, text=' ', bg=window_background_color, font=fonts["big bold"])
        self.l2_bis = Label(self, text=' ' * 5, bg=window_background_color, font=fonts["big bold"])
        self.l3 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l4 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l5 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l6 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        
        self.l0.grid(row=0, columnspan=5)
        self.l1.grid(row=2, columnspan=5)
        self.l2.grid(row=3, column=1, rowspan=7)
        self.l2_bis.grid(row=3, column=3, rowspan=7)
        self.l3.grid(row=4, column=4)
        self.l4.grid(row=6, column=4)
        self.l5.grid(row=11, columnspan=5)
        self.l6.grid(row=12, columnspan=5)
        self.widgets['l0'] = self.l0
        self.widgets['l1'] = self.l1
        self.widgets['l2'] = self.l2
        self.widgets['l2_bis'] = self.l2_bis
        self.widgets['l3'] = self.l3
        self.widgets['l4'] = self.l4
        self.widgets['l5'] = self.l5
        self.widgets['l6'] = self.l6
        
        self.pdf_list = Listbox(
            self,
            listvariable=StringVar(value=[]),
            font=fonts["normal"],
            fg=font_color,
            width=60,
            height=15,
        )
        self.scrollbarv = ttk.Scrollbar(self, orient=VERTICAL, command=self.pdf_list.yview)
        self.scrollbarh = ttk.Scrollbar(self, orient=HORIZONTAL, command=self.pdf_list.xview)
        self.pdf_list.configure(yscrollcommand=self.scrollbarv.set, xscrollcommand=self.scrollbarh.set)
        self.pdf_list.bind("<<ListboxSelect>>", self.file_select)
        self.pdf_list.bind("<Double-1>", self.view_pdf)
        self.pdf_list.bind("<Return>", self.view_pdf)
        self.pdf_list.grid(row=3, column=0, rowspan=6, sticky=(N, W, E, S))
        self.scrollbarv.grid(row=3, column=0, rowspan=6, sticky=(N, E, S))
        self.scrollbarh.grid(row=10, column=0, sticky=(W, E, S))
        self.widgets['pdf_list'] = self.pdf_list
        self.widgets['scrollbarv'] = self.scrollbarv
        self.widgets['scrollbarh'] = self.scrollbarh
        
        up_arrow_img = Image.open(self.params['up_arrow_img_path'])
        up_arrow_img = up_arrow_img.resize((20, 50), Image.ANTIALIAS)
        down_arrow_img = up_arrow_img.transpose(Image.FLIP_TOP_BOTTOM)
        self.up_arrow_img = ImageTk.PhotoImage(up_arrow_img)
        self.down_arrow_img = ImageTk.PhotoImage(down_arrow_img)
        
        self.up_arrow_button = Button(
            self,
            image=self.up_arrow_img,
            command=self.move_up,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
        )
        self.up_arrow_button.grid(row=3, column=2, rowspan=2)
        self.widgets['uo_arrow_button'] = self.up_arrow_button
        
        self.down_arrow_button = Button(
            self,
            image=self.down_arrow_img,
            command=self.move_down,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
        )
        self.down_arrow_button.grid(row=7, column=2, rowspan=2)
        self.widgets['down_arrow_button'] = self.down_arrow_button
        
        self.load_button = Button(
            self,
            text=texts[self.params['language']]['open_file'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=30,
            command=partial(self.load_files, True, [('PDF', '*.pdf')]), #, ('Images', '*.jpg *.jpeg *.JPG *.JPEG *.png'), ('PDF and Images', '*.pdf *.jpg *.jpeg *.JPG *.JPEG *.png')]),
        )
        self.load_button.grid(row=3, column=4)
        self.widgets['load_button'] = self.load_button
        
        self.merge_button = Button(
            self,
            text=texts[self.params['language']]['page2_merge_button'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=30,
            command=self.merge_files,
        )
        self.merge_button.grid(row=5, column=4)
        self.widgets['merge_button'] = self.merge_button
        
        self.clear_button = Button(
            self,
            text=texts[self.params['language']]['page2_clear_button'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=30,
            command=self.clear_pdfs,
        )
        self.clear_button.grid(row=7, column=4)
        self.widgets['clear_button'] = self.clear_button
        
        self.go_back_button = Button(
            self,
            text=texts[self.params['language']]['go_back'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=30,
            command=self.form_start_page,
        )
        self.go_back_button.grid(row=13, column=0, columnspan=5)
        self.widgets['go_back_button'] = self.go_back_button
    
    
    def load_split_page(self):
        
        self.erase_page()
        self.parent.title(texts[self.params['language']]['page3_suptitle'])
        
        self.filepath = None
        self.start_page = None
        self.stop_page = None
        
        self.title = Label(
            self,
            text=texts[self.params['language']]['page3_title'],
            font=fonts["title"],
            fg=font_color,
            bg=window_background_color,
        )
        self.title.grid(row=1, columnspan=10)
        self.widgets['title'] = self.title
        
        self.load_button = Button(
            self,
            text=texts[self.params['language']]['open_file'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            width=30,
            command=partial(self.load_file, False, [('PDF', '*.pdf')]),
            cursor='hand2',
        )
        self.load_button.grid(row=3, column=8, columnspan=2)
        self.widgets['load_button'] = self.load_button
        
        self.start_page_label = Label(
            self,
            text = texts[self.params['language']]['page3_start_page'],
            anchor='w',
            bg=window_background_color,
            font=fonts["bold"],
            fg=font_color,
            width=15,
        )
        self.start_page_label.grid(row=5, column=8, sticky=(W, E))
        self.widgets['start_page_label'] = self.start_page_label
        
        self.stop_page_label = Label(
            self,
            text = texts[self.params['language']]['page3_stop_page'],
            anchor='w',
            bg=window_background_color,
            font=fonts["bold"],
            fg=font_color,
            width=15,
        )
        self.stop_page_label.grid(row=6, column=8, sticky=(W, E))
        self.widgets['stop_page_label'] = self.stop_page_label
        
        self.start_page_sv = StringVar()
        self.start_page_sv.trace("w", lambda name, index, mode, sv=self.start_page_sv: self.retrieve_page_number("start"))
        self.start_page_entry = Entry(
            self,
            bg = entry_bg_color,
            font=fonts["normal"],
            fg=font_color,
            width=15,
            textvariable=self.start_page_sv,
        )
        self.start_page_entry.grid(row=5, column=9, sticky=(W, E))
        self.widgets['start_page_entry'] = self.start_page_entry
        
        self.stop_page_sv = StringVar()
        self.stop_page_sv.trace("w", lambda name, index, mode, sv=self.stop_page_sv: self.retrieve_page_number("stop"))
        self.stop_page_entry = Entry(
            self,
            bg = entry_bg_color,
            font=fonts["normal"],
            fg=font_color,
            width=15,
            textvariable=self.stop_page_sv,
        )
        self.stop_page_entry.grid(row=6, column=9, sticky=(W, E))
        self.widgets['stop_page_entry'] = self.stop_page_entry
        
        self.split_button = Button(
            self,
            text=texts[self.params['language']]['page3_split_button'],
            font=fonts['bold'],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            width=30,
            command=self.split_pdf,
            cursor='hand2',
        )
        self.split_button.grid(row=8, column=8, columnspan=2)
        self.widgets['split_button'] = self.split_button
        
        self.pdf_text = Text(
            self,
            width=20,
            height=10,
            font=fonts['title'],
            fg=font_color,
        )
        self.scrollbarv = ttk.Scrollbar(self, orient=VERTICAL, command=self.pdf_text.yview)
        self.scrollbarh = ttk.Scrollbar(self, orient=HORIZONTAL, command=self.pdf_text.xview)
        self.pdf_text.configure(yscrollcommand=self.scrollbarv.set, xscrollcommand=self.scrollbarh.set)
        self.pdf_text.grid(row=3, column=0, rowspan=6, columnspan=7, sticky=(N, W, E, S))
        self.scrollbarv.grid(row=3, column=6, rowspan=6, sticky=(N, E, S))
        self.scrollbarh.grid(row=8, column=0, columnspan=7, sticky=(W, E, S))
        
        self.zoom_label = Label(
            self,
            text=texts[self.params['language']]['page3_zoom_label'],
            font=fonts['bold'],
            fg=font_color,
            bg=window_background_color,
            anchor='w',
        )
        self.zoom_label.grid(row=9, column=0)
        self.widgets['zoom_label'] = self.zoom_label
        
        self.zoom_scale = Scale(
            self,
            from_=50,
            to=500,
            resolution=5,
            orient=HORIZONTAL,
            font=fonts['normal'],
            fg=font_color,
            borderwidth=0,
            bg=window_background_color,
            activebackground=active_background_color,
            command=self.change_zoom,
            cursor='hand2',
        )
        self.zoom_scale.set(100)
        self.zoom_scale.grid(row=9, column=1, columnspan=6, sticky=(W, E))
        self.widgets['zoom_scale'] = self.zoom_scale
        
        left_arrow_img = Image.open(self.params['up_arrow_img_path'])
        left_arrow_img = left_arrow_img.resize((20, 50), Image.ANTIALIAS)
        right_arrow_img = left_arrow_img.transpose(Image.Transpose.ROTATE_270)
        left_arrow_img = left_arrow_img.transpose(Image.Transpose.ROTATE_90)
        self.left_arrow_img = ImageTk.PhotoImage(left_arrow_img)
        self.right_arrow_img = ImageTk.PhotoImage(right_arrow_img)
        
        self.change_page_label = Label(
            self,
            text=texts[self.params['language']]['page3_change_page_label'],
            font=fonts['bold'],
            fg=font_color,
            bg=window_background_color,
            anchor='w',
        )
        self.change_page_label.grid(row=11, column=0)
        self.widgets['change_page_label'] = self.change_page_label
        
        self.left_arrow_button = Button(
            self,
            image=self.left_arrow_img,
            command=self.decrease_page_nb,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=40,
        )
        self.left_arrow_button.grid(row=11, column=1, sticky=(W,))
        self.widgets['left_arrow_button'] = self.left_arrow_button
        
        self.right_arrow_button = Button(
            self,
            image=self.right_arrow_img,
            command=self.increase_page_nb,
            bg=button_background_color,
            activebackground=active_background_color,
            cursor='hand2',
            width=40,
        )
        self.right_arrow_button.grid(row=11, column=6, sticky=(E,))
        self.widgets['right_arrow_button'] = self.right_arrow_button
        
        self.page_nb_sv = StringVar()
        self.page_nb_sv.trace("w", lambda name, index, mode, sv=self.page_nb_sv: self.retrieve_page_number("display"))
        self.page_nb_entry = Entry(
            self,
            bg = entry_bg_color,
            font=fonts["normal"],
            fg=font_color,
            width=8,
            textvariable=self.page_nb_sv,
        )
        self.page_nb_entry.grid(row=11, column=3)
        self.current_page = 1
        self.page_nb_sv.set(self.current_page)
        self.widgets['page_nb_entry'] = self.page_nb_entry
        
        self.total_page_nb_sv = StringVar()
        self.total_page_nb_sv.set('/')
        self.total_page_nb_label = Label(
            self,
            textvariable=self.total_page_nb_sv,
            width=8,
            bg=window_background_color,
            font=fonts["normal"],
            fg=font_color,
        )
        self.total_page_nb_label.grid(row=11, column=4)
        self.widgets['total_page_nb_label'] = self.total_page_nb_label
        
        self.go_back_button = Button(
            self,
            text=texts[self.params['language']]['go_back'],
            font=fonts["bold"],
            fg=font_color,
            bg=button_background_color,
            activebackground=active_background_color,
            width=30,
            command=self.form_start_page,
            cursor='hand2',
        )
        self.go_back_button.grid(row=14, column=0, columnspan=10)
        self.widgets['go_back_button'] = self.go_back_button
        
        max_length = np.max([len(texts[self.params['language']][key]) for key in texts[self.params['language']].keys() if 'page2' in key])
        self.l0 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l1 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l2 = Label(self, text=' ' * 5, bg=window_background_color, font=fonts["big bold"])
        self.l3 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l4 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l5 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l6 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l7 = Label(self, text=' ' * max_length, bg=window_background_color, font=fonts["big bold"])
        self.l8 = Label(self, text=' ', bg=window_background_color, font=fonts["big bold"], width=2)
        self.l9 = Label(self, text=' ', bg=window_background_color, font=fonts["big bold"], width=2)
        
        self.l0.grid(row=0, columnspan=10)
        self.l1.grid(row=2, columnspan=10)
        self.l2.grid(row=3, column=7, rowspan=6)
        self.l3.grid(row=4, column=8, columnspan=2)
        self.l4.grid(row=7, column=8, columnspan=2)
        self.l5.grid(row=10, columnspan=10)
        self.l6.grid(row=12, columnspan=10)
        self.l7.grid(row=13, columnspan=10)
        self.l8.grid(row=11, column=2)
        self.l9.grid(row=11, column=5)
        self.widgets['l0'] = self.l0
        self.widgets['l1'] = self.l1
        self.widgets['l2'] = self.l2
        self.widgets['l3'] = self.l3
        self.widgets['l4'] = self.l4
        self.widgets['l5'] = self.l5
        self.widgets['l6'] = self.l6
        self.widgets['l7'] = self.l7
        self.widgets['l8'] = self.l8
        self.widgets['l9'] = self.l9
        
        self.widgets['pdf_text'] = self.pdf_text
        self.widgets['scrollbarv'] = self.scrollbarv
        self.widgets['scrollbarh'] = self.scrollbarh
    
    
    def stop_before_complete(self):
        resp = messagebox.askyesno("Attention","Vous allez perdre les données déjà saisies.\nVoulez vous continuer malgré tout ?", icon='warning')
        if resp:
            self.quit()
    
    def cancel_before_complete(self):
        resp = messagebox.askyesno("Attention","Vous allez perdre les données déjà saisies.\nVoulez vous revenir au début malgré tout ?", icon='warning')
        if resp:
            self.data = {}
            self.form_start_page()