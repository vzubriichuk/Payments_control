# -*- coding: utf-8 -*-
"""
Created on Wed May 15 22:51:04 2019

@author: v.shkaberda
"""

from datetime import datetime
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from math import ceil
import tkinter as tk

# example of subsription and default recipient
EMAIL_TO = b'\xd0\xa4\xd0\xbe\xd0\xb7\xd0\xb7\xd0\xb8|\
\xd0\x9b\xd0\xbe\xd0\xb3\xd0\xb8\xd1\x81\xd1\x82\xd0\xb8\xd0\xba\xd0\xb0|\
\xd0\x90\xd0\xbd\xd0\xb0\xd0\xbb\xd0\xb8\xd1\x82\xd0\xb8\xd0\xba\xd0\xb8'.decode()


class LoginError(tk.Tk):
    ''' Raise an error when user doesn't have permission to work with db.
    '''
    def __init__(self):
        super().__init__()

        # Do not show main window
        self.withdraw()
        messagebox.showerror(
                'Ошибка подключения',
                'Нет прав для работы с сервером.\n'
                'Обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class AccessError(tk.Tk):
    ''' Raise an error when user doesn't have permission to work with app.
    '''
    def __init__(self):
        super().__init__()

        # Do not show main window
        #self.withdraw()
        messagebox.showerror(
            'Ошибка доступа',
            'Нет прав для работы с приложением.\n'
            'Для получения прав обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class PaymentApp(tk.Tk):
    def __init__(self, **kwargs):
        super().__init__()
        self.title('Платежи')
        # geometry_storage {Framename:(width, height)}
        self._geometry = {'MainMenu': (250, 320),
                          'CreateForm': (880, 600),
                          'PreviewForm': (1200, 600),
                          'DiscardForm': (880, 400),
                          'ApproveForm': (1200, 350)}
        # Customize header style (used in PreviewForm)
        style = ttk.Style()
        try:
            style.element_create("LightBlue.Treeheading.border", "from", "default")
            style.layout("LightBlue.Treeview.Heading", [
                ("LightBlue.Treeheading.cell", {'sticky': 'nswe'}),
                ("LightBlue.Treeheading.border", {'sticky':'nswe', 'children': [
                    ("Custom.Treeheading.padding", {'sticky':'nswe', 'children': [
                        ("Custom.Treeheading.image", {'side':'right', 'sticky':''}),
                        ("Custom.Treeheading.text", {'sticky':'we'})
                    ]})
                ]}),
            ])
            style.configure("LightBlue.Treeview.Heading",
                background="lightblue", foreground="black", relief='groove', font=('Calibri', 10, 'bold'))
            style.map("LightBlue.Treeview.Heading",
                relief=[('active','sunken'),('pressed','flat')])

            style.map('ButtonGreen.TButton')
            style.configure('ButtonGreen.TButton', foreground='green')

            style.map('ButtonRed.TButton')
            style.configure('ButtonRed.TButton', foreground='red')

            style.map('ButtonMenu.TButton')
            style.configure('ButtonMenu.TButton',
                            font=('Calibri', 16),
                            borderwidth='4')

            style.configure("Big.TLabelframe.Label", font=("Calibri", 14))
            style.configure("TMenubutton", background='white')
        except tk.TclError:
            # if during debug previous tk wasn't destroyed and style remains modified
            pass

        # the container is where we'll stack a bunch of frames
        # the one we want to be visible will be raised above others
        container = tk.Frame(self)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._frames = {}
        for F in (MainMenu, CreateForm, PreviewForm, DiscardForm, ApproveForm):
            frame_name = F.__name__
            frame = F(parent=container, controller=self, **kwargs)
            self._frames[frame_name] = frame
            # put all of them in the same location
            frame.grid(row=0, column=0, sticky='nsew')

        self._show_frame('MainMenu')

    def _center_window(self, w, h):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width/2) - (w/2))
        start_y = int((screen_height/2) - (h/2))

        self.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))

    def _show_frame(self, frame_name):
        ''' Show a frame for the given frame name
        '''
        if frame_name == 'MainMenu':
            self.resizable(width=False, height=False)
        else:
            self.resizable(width=True, height=True)
        frame = self._frames[frame_name]
        frame.tkraise()
        self._center_window(*(self._geometry[frame_name]))
        if frame_name in ('DiscardForm', 'ApproveForm'):
            frame._refresh()
        if frame_name in ('PreviewForm', 'ApproveForm'):
            frame._resize_columns()


class PaymentFrame(tk.Frame):
    def __init__(self, parent, controller, connection, shortusername, userID):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        self._shortusername = shortusername
        self.userID = userID
        self.conn = connection

    def _add_user_label(self, parent):
        user_label = tk.Label(parent, text=self._shortusername, padx=10)
        user_label.pack(side=tk.RIGHT, anchor=tk.NE)

    def _validate_sum(self, sum_entry):
        ''' Validation of self.sum_entry'''
        sum_entry = sum_entry.replace(',', '.')
        if not sum_entry:
            return True
        try:
            if float(sum_entry) < 10**9:
                return True
        except (TypeError, ValueError):
            return False


class MainMenu(PaymentFrame):
    def __init__(self, parent, controller, connection, shortusername, userID, **kwargs):
        #tk.Frame.__init__(self, parent)
        #self._shortusername = shortusername
        super().__init__(parent, controller, connection, shortusername, userID)

        top = tk.Frame(self, name='top', padx=5)
        top.pack(side=tk.TOP, fill=tk.X)

        # Label to customize bottom text
        bottommenu = ttk.LabelFrame(self, text='Выберите действие:', name='bottommenu',
                                    style='Big.TLabelframe'
                                )
        bottommenu.pack(side=tk.TOP, fill=tk.Y, expand=False, padx=10, pady=5)

        self._add_user_label(top)

        bt1 = ttk.Button(bottommenu, text="Создать", width=17, style='ButtonMenu.TButton',
                         command=lambda: controller._show_frame('CreateForm'))
        bt1.pack(side=tk.TOP, padx=15, pady=7, expand=False)

        bt2 = ttk.Button(bottommenu, text="Отменить", width=17, style='ButtonMenu.TButton',
                         command=lambda: controller._show_frame('DiscardForm'))
        bt2.pack(side=tk.TOP, padx=15, pady=7, expand=False)

        bt3 = ttk.Button(bottommenu, text="Утвердить", width=17, style='ButtonMenu.TButton',
                         command=lambda: controller._show_frame('ApproveForm'))
        bt3.pack(side=tk.TOP, padx=15, pady=7, expand=False)

        bt4 = ttk.Button(bottommenu, text="Просмотр", width=17, style='ButtonMenu.TButton',
                         command=lambda: controller._show_frame('PreviewForm'))
        bt4.pack(side=tk.TOP, padx=15, pady=7, expand=False)

        bt5 = ttk.Button(bottommenu, text="Выход", width=17, style='ButtonMenu.TButton',
                         command=controller.destroy)
        bt5.pack(side=tk.TOP, padx=15, pady=7, expand=False)


class CreateForm(PaymentFrame):
    def __init__(self, parent, controller, connection, userID, shortusername, mvz, **kwargs):
        super().__init__(parent, controller, connection, shortusername, userID)
        #self.mvznames, self.mvzSAP = zip(*mvz)
        self.mvz = dict(mvz)

        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)

        main_label = tk.Label(top, text='Форма создания заявки на согласование',
                              padx=10, font=('Calibri', 10, 'bold'))
        main_label.pack(side=tk.LEFT, expand=False, anchor=tk.NW)

        self._add_user_label(top)

        # First Fill Frame with (MVZ, office, contragentID)
        row1_cf = tk.Frame(self, name='row1_cf', padx=15)

        self.mvz_label = tk.Label(row1_cf, text='МВЗ', padx=10)

        self.mvz_current = tk.StringVar()
        #self.mvz_current.set(self.mvznames[0]) # default value
        self.mvz_box = ttk.OptionMenu(row1_cf, self.mvz_current, '', *self.mvz.keys(),
                                      command = self._choose_mvz)
        self.mvz_box.config(width=35)

#        self.mvz_box = ttk.Combobox(row1_cf, width=35)
#        self.mvz_box['values'] = self.mvznames
#        self.mvz_box.bind("<<ComboboxSelected>>", self._choose_mvz)

        self.mvz_sap = tk.Label(row1_cf, padx=5, bg='lightgray', width=10)

        self.office_label = tk.Label(row1_cf, text='Офис', padx=20)
        self.office_box = ttk.Combobox(row1_cf, width=20)
        #self.office_box['values'] = self.mvzSAP
        self.office_box['values'] = list(self.mvz.values())

        self.contragent_label = tk.Label(row1_cf, text='Контрагент', padx=20)
        self.contragent_entry = tk.Entry(row1_cf, width=21)

        # Pack row1_cf
        self._row1_pack()

        # Second Fill Frame with (Plan date, Sum, Tax)
        row2_cf = tk.Frame(self, name='row2_cf', padx=15)

        self.plan_date_label = tk.Label(row2_cf, text='Плановая дата', padx=10)
        self.plan_date_entry = DateEntry(row2_cf, width=12, state='readonly',
                    font=('Calibri', 10), selectmode='day', borderwidth=2)

        self.sum_label = tk.Label(row2_cf, text='Сумма без НДС', padx=20)
        self.sumtotal = tk.DoubleVar()
        self.sumtotal.set('0.00')
        vcmd = (self.register(self._validate_sum))
        self.sum_entry = tk.Entry(row2_cf, width=20, textvariable=self.sumtotal,
                        validate='all', validatecommand=(vcmd, '%P'))
        # Alternative validation
        #self.sum_entry = tk.Entry(row2_cf, width=20, textvariable=self.sumtotal)
        #self.sum_entry.bind("<FocusOut>", self._on_focus_out)

        self.nds_label = tk.Label(row2_cf, text='НДС', padx=20)
        self.nds = tk.IntVar()
        self.nds.set(20)
        self.nds20 = ttk.Radiobutton(row2_cf, text="20 %", variable=self.nds, value=20)
        self.nds7 = ttk.Radiobutton(row2_cf, text="7 %", variable=self.nds, value=7)
        self.nds0 = ttk.Radiobutton(row2_cf, text="0 %", variable=self.nds, value=0)

        # Pack row2_cf
        self._row2_pack()

        # Text Frame
        text_cf = ttk.LabelFrame(self, text=' Описание заявки ', name='text_cf')

        self.desc_text = tk.Text(text_cf)    # input and output box
        self.desc_text.pack(in_=text_cf, expand=True, pady=15)

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt3 = ttk.Button(bottom_cf, text="Меню", width=10,
                         command=lambda: controller._show_frame('MainMenu'))
        bt3.pack(side=tk.RIGHT, padx=15, pady=5)

        bt2 = ttk.Button(bottom_cf, text="Очистить", width=10,
                         command=self._clear, style='ButtonRed.TButton')
        bt2.pack(side=tk.RIGHT, padx=15, pady=5)

        bt1 = ttk.Button(bottom_cf, text="Создать", width=10,
                         command=self._create_request, style='ButtonGreen.TButton')
        bt1.pack(side=tk.RIGHT, padx=15, pady=5)

        # Pack frames
        top.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X)
        row1_cf.pack(side=tk.TOP, fill=tk.X)
        row2_cf.pack(side=tk.TOP, fill=tk.X)
        text_cf.pack(side=tk.TOP, fill=tk.X, expand=True, padx=10, pady=5)

    def _choose_mvz(self, event):
        #self.mvz_sap.config(text=self.mvzSAP[self.mvz_box.current()])
        self.mvz_sap.config(text=self.mvz[self.mvz_current.get()])

    def _clear(self):
        #self.mvz_box.set('')
        self.mvz_current.set('')
        self.mvz_sap.config(text='')
        self.office_box.set('')
        self.contragent_entry.delete(0, tk.END)
        self.plan_date_entry.delete(0, tk.END)
        self.sumtotal.set('0.00')
        self.nds.set(20)
        self.desc_text.delete("1.0", tk.END)

#    def _on_focus_out(self, event):
#        ''' Function to be bind to focus out of self.sum_entry'''
#        sum_entry = self.sum_entry.get().replace(',', '.')
#        try:
#            self.sumtotal.set('{:.2f}'.format(float(sum_entry)))
#        except (TypeError, ValueError):
#            self.sumtotal.set('0.00')
#            messagebox.showerror(
#                    'Некорректная сумма',
#                    'Введена некорретная сумма!'
#                    )

    def _validate_plan_date(self):
        date = self.plan_date_entry.get()
        date = datetime.strptime(date, '%d.%m.%Y')
        today = datetime.now()
        return date > today

    def _create_request(self):
        if not self.mvz_current.get():
            messagebox.showerror(
                    'Создание заявки',
                    'Не указано МВЗ'
            )
            return
        if not self._validate_plan_date():
            messagebox.showerror(
                    'Создание заявки',
                    'Дата не может быть сегодняшней или ранее'
            )
            return
        request = {'mvz': self.mvz_sap.cget('text'),
                   'office': self.office_box.current(),
                   'contragent': self.contragent_entry.get() or None,
                   'plan_date': self.plan_date_entry.get(),
                   'sumtotal': self.sumtotal.get() if self.sum_entry.get() else 0.,
                   'nds':  self.nds.get(),
                   'text': self.desc_text.get("1.0", tk.END)
                }
        created_success = self.conn.create_request(userID=self.userID, **request)
        if created_success:
            messagebox.showinfo(
                    'Создание заявки',
                    'Заявка создана'
            )
            self._clear()
            self.controller._show_frame('MainMenu')
        else:
            messagebox.showerror(
                    'Создание заявки',
                    'Произошла ошибка при создании заявки'
            )

    def _row1_pack(self):
        self.mvz_label.pack(side=tk.LEFT)
        self.mvz_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.mvz_sap.pack(side=tk.LEFT)
        self.contragent_entry.pack(side=tk.RIGHT, padx=10, pady=5)
        self.contragent_label.pack(side=tk.RIGHT)
        self.office_box.pack(side=tk.RIGHT, padx=5, pady=5)
        self.office_label.pack(side=tk.RIGHT)

    def _row2_pack(self):
        self.plan_date_label.pack(side=tk.LEFT)
        self.plan_date_entry.pack(side=tk.LEFT, padx=5, pady=5)
        self.nds0.pack(side=tk.RIGHT, padx=7)
        self.nds7.pack(side=tk.RIGHT, padx=7)
        self.nds20.pack(side=tk.RIGHT, padx=8)
        self.nds_label.pack(side=tk.RIGHT, padx=5)
        self.sum_entry.pack(side=tk.RIGHT, padx=5, pady=5)
        self.sum_label.pack(side=tk.RIGHT)


class PreviewForm(PaymentFrame):
    def __init__(self, parent, controller, connection, userID, shortusername, mvz, **kwargs):
        super().__init__(parent, controller, connection, shortusername, userID)
        self.approveform_bool = isinstance(self, ApproveForm)
        self.mvznames, _ = zip(*mvz)

        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)

        main_label = tk.Label(top, text='На утверждении' if self.approveform_bool else 'Просмотр заявок',
                              padx=10, font=('Calibri', 10, 'bold'))
        main_label.pack(side=tk.LEFT, expand=False, anchor=tk.NW)

        self._add_user_label(top)

        top.pack(side=tk.TOP, fill=tk.X, expand=False)

        # Filters for PreviewForm
        if not self.approveform_bool:
            filterframe = ttk.LabelFrame(self, text=' Фильтры ', name='filterframe')

            # First Filter Frame with (MVZ, office, contragentID)
            row1_cf = tk.Frame(filterframe, name='row1_cf', padx=15)

            self.mvz_label = tk.Label(row1_cf, text='МВЗ', padx=10)

            self.mvz_current = tk.StringVar()
            self.mvz_current.set(self.mvznames[0]) # default value

            self.mvz_box = ttk.Combobox(row1_cf, width=35)
            self.mvz_box['values'] = self.mvznames

            self.office_label = tk.Label(row1_cf, text='Офис', padx=20)
            self.office_box = ttk.Combobox(row1_cf, width=20)
            self.office_box['values'] = self.mvznames

            self.contragent_label = tk.Label(row1_cf, text='Контрагент', padx=20)
            self.contragent_entry = tk.Entry(row1_cf, width=21)

            # Pack row1_cf
            self._row1_pack()
            row1_cf.pack(side=tk.TOP, fill=tk.X)

            # Second Fill Frame with (Plan date, Sum, Tax)
            row2_cf = tk.Frame(filterframe, name='row2_cf', padx=15)

            self.plan_date_label_from = tk.Label(row2_cf, text='Плановая дата: от', padx=10)
            self.plan_date_entry_from = DateEntry(row2_cf, width=12,
                        font=('Calibri', 10), selectmode='day', borderwidth=2)
            self.plan_date_label_to = tk.Label(row2_cf, text='до', padx=10)

            self.plan_date_entry_3 = tk.Label(row2_cf, text='Плановая дата: 2', padx=10)

            self.cal2 = tk.Frame(row2_cf, name='row2_cf', padx=15)
            self.plan_date_entry_to = DateEntry(self.cal2, width=12,
                        font=('Calibri', 10), selectmode='day', borderwidth=2)

            self.sum_label = tk.Label(row2_cf, text='Сумма без НДС', padx=20)
            self.sumtotal = tk.DoubleVar()
            self.sumtotal.set('0.00')
            vcmd = (self.register(self._validate_sum))
            self.sum_entry = tk.Entry(row2_cf, width=20, textvariable=self.sumtotal,
                            validate='all', validatecommand=(vcmd, '%P'))
            # Alternative validation
            #self.sum_entry = tk.Entry(row2_cf, width=20, textvariable=self.sumtotal)
            #self.sum_entry.bind("<FocusOut>", self._on_focus_out)

            self.nds_label = tk.Label(row2_cf, text='НДС', padx=20)
            self.nds = tk.IntVar()
            self.nds.set(20)
            self.nds20 = ttk.Radiobutton(row2_cf, text="20 %", variable=self.nds, value=20)
            self.nds7 = ttk.Radiobutton(row2_cf, text="7 %", variable=self.nds, value=7)
            self.nds0 = ttk.Radiobutton(row2_cf, text="0 %", variable=self.nds, value=0)

            # Pack row1_cf
            self._row2_pack()
            row2_cf.pack(side=tk.TOP, fill=tk.X)
            filterframe.pack(side=tk.TOP, fill=tk.X, expand=False, padx=10, pady=5)

        # Text Frame
        preview_cf = ttk.LabelFrame(self, text=' Заявки ', name='preview_cf')

        # column name and width
        #self.headings=('a', 'bb', 'cccc')  # for debug
        self.headings = {'№':10, 'Инициатор':140, 'Дата создания':100, 'МВЗ':60,
                    'Офис':100, 'Контрагент':60, 'Плановая дата':100,
                    'Сумма без НДС':100, 'Сумма с НДС':100, 'Статус':30,
                    'Описание':120, 'Краткое описание':120}

        self.table = ttk.Treeview(preview_cf, show="headings", selectmode="browse", style="LightBlue.Treeview")
        self._init_table(preview_cf)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt3 = ttk.Button(bottom_cf, text="Меню", width=10, command=lambda: controller._show_frame('MainMenu'))
        bt3.pack(side=tk.RIGHT, padx=15, pady=5)

        bt2 = ttk.Button(bottom_cf, text="Подробно", width=10,
                         command=self._show_detail)
        bt2.pack(side=tk.RIGHT, padx=15, pady=5)

        if not self.approveform_bool:
            bt1 = ttk.Button(bottom_cf, text="Обновить", width=10,
                             command=self._refresh)
            bt1.pack(side=tk.RIGHT, padx=15, pady=5)

        # Pack frames
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        preview_cf.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _init_table(self, parent):
        if isinstance(self.headings, dict):
            self.table["columns"] = tuple(self.headings.keys())
            self.table["displaycolumns"] = tuple(k for k in self.headings.keys()
                if k not in ('НДС', 'Описание') and not (self.approveform_bool and k == 'Статус'))
            for head, width in self.headings.items():
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=width, anchor=tk.CENTER)

        else:
            self.table["columns"] = self.headings
            self.table["displaycolumns"] = self.headings
            for head in self.headings:
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=50*len(head), anchor=tk.CENTER)

        #self._show_rows(rows=((123, 456, 789), ('abc', 'def', 'ghk')))  # for debug

        self.table.tag_configure('Отм.', background='lightgray')
        self.table.tag_configure('Утв.', background='lightgreen')
        self.table.tag_configure('Откл.', background='#f66e6e')
        #self.table.tag_configure('oddrow', background='lightgray')

        self.table.bind("<Double-1>", self._show_detail)

        scrolltable = tk.Scrollbar(parent, command=self.table.yview)
        self.table.configure(yscrollcommand=scrolltable.set)
        scrolltable.pack(side=tk.RIGHT, fill=tk.Y)

    def _resize_columns(self):
        for head, width in self.headings.items():
            self.table.column(head, width=width)

    def _row1_pack(self):
        self.mvz_label.pack(side=tk.LEFT)
        self.mvz_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.contragent_entry.pack(side=tk.RIGHT, padx=10, pady=5)
        self.contragent_label.pack(side=tk.RIGHT)
        self.office_box.pack(side=tk.RIGHT, padx=5, pady=5)
        self.office_label.pack(side=tk.RIGHT)

    def _row2_pack(self):
        self.plan_date_label_from.pack(side=tk.LEFT)
        self.plan_date_entry_from.pack(side=tk.LEFT, padx=5, pady=5)
        self.plan_date_label_to.pack(side=tk.LEFT)
        self.cal2.pack(side=tk.LEFT)
        self.plan_date_entry_to.pack()
        self.plan_date_entry_3.pack(side=tk.LEFT)
        self.nds0.pack(side=tk.RIGHT, padx=7)
        self.nds7.pack(side=tk.RIGHT, padx=7)
        self.nds20.pack(side=tk.RIGHT, padx=8)
        self.nds_label.pack(side=tk.RIGHT, padx=5)
        self.sum_entry.pack(side=tk.RIGHT, padx=5, pady=5)
        self.sum_label.pack(side=tk.RIGHT)

    def _refresh(self):
        ''' Extract information from filters '''
        rows = self.conn.get_paymentslist()
        self._show_rows(rows)

    def _show_detail(self, event=None):
        ''' Show details when double-clicked on row '''
        if not event or event.y > 19:
            curRow = self.table.focus()
            if curRow:
                newlevel = tk.Toplevel(self.parent)
                newlevel.title('Заявка детально')
                if self.approveform_bool:
                    ApproveConfirmation(newlevel, self, self.conn, self.userID,
                                    self.headings,
                                    self.table.item(curRow).get('values'))
                else:
                    DetailedPreview(newlevel, self, self.conn, self.userID,
                                    self.headings,
                                    self.table.item(curRow).get('values'))
                newlevel.resizable(width=False, height=False)
                newlevel.focus()

    def _show_rows(self, rows=((777, 88, 9),
                                 ('abce', 'de`24f', 'gh`24k'),
                                 ('', '1', ''),
                                 ('123123', '11', '2'))*3):
        ''' Refresh table with new rows '''
        self.table.delete(*self.table.get_children())

        for row in rows:
            # tag = (Status,)
            self.table.insert('', tk.END,
                              values=tuple(row),
                              tags=(row[-3],))


class DiscardForm(PaymentFrame):
    def __init__(self, parent, controller, connection, userID, shortusername, **kwargs):
        super().__init__(parent, controller, connection, shortusername, userID)

        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)

        main_label = tk.Label(top, text='Отмена заявок\nВыделите заявки, '
         'которые необходимо отменить, и нажмите кнопку "Отменить"',
         padx=10, justify=tk.LEFT, font=('Calibri', 10, 'bold'))
        main_label.pack(side=tk.LEFT, expand=False, anchor=tk.NW)

        self._add_user_label(top)

        # Middle Frame
        discard_f = ttk.LabelFrame(self, text=' Заявки на согласовании ', name='discard_f')

        self.discardbox = tk.Listbox(discard_f, selectmode='multiple')
        self.discardbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.discardID = None  # Name to store ID for discarding

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt2 = ttk.Button(bottom_cf, text="Меню", width=10,
                         command=lambda: controller._show_frame('MainMenu'))
        bt2.pack(side=tk.RIGHT, padx=15, pady=5)

        bt1 = ttk.Button(bottom_cf, text="Отменить", width=10,
                         command=self._discard_request)
        bt1.pack(side=tk.RIGHT, padx=15, pady=5)

        # Pack frames
        top.pack(side=tk.TOP, fill=tk.X, expand=False)
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        discard_f.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _discard_request(self):
        mboxname = 'Отмена заявки'
        discardset = self.discardbox.curselection()
        if not discardset:
            messagebox.showerror(mboxname, 'Выберите хотя бы одну заявку')
            return
        confirmed = messagebox.askyesno(title=mboxname,
                  message='Вы собираетесь отменить заявок: {}\nПродолжить?'.format(len(discardset)))
        if confirmed:
            self.conn.update_discarded(tuple(self.discardID[i] for i in discardset))
            messagebox.showinfo(
                    mboxname,
                    'Заявок отменено: {}'.format(len(discardset))
            )
            self._refresh()
        else:
            messagebox.showinfo(
                    mboxname,
                    'Действие отменено'
            )

    def _get_list_to_discard(self):
        discardlist = self.conn.get_discardlist(self.userID)
        self.discardID = [d[0] for d in discardlist]
        discardlist = tuple(
                map(lambda d: 'Заявка №{} от {}, МВЗ: {}, офис: {}, '
                    'план.дата: {}, сумма: {}, описание: {}'.format(*d),
                    discardlist
                ))
        return discardlist

    def _refresh(self):
        self.discardbox.delete(0, 'end')
        self.discardbox.insert('end', *self._get_list_to_discard())


class ApproveForm(PreviewForm):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def _refresh(self):
        ''' Refresh  '''
        self.table.delete(*self.table.get_children())
        rows = self.conn.get_approvelist(self.userID)
        self._show_rows(rows)


class DetailedRequest(tk.Frame):
    def __init__(self, parent, parentform, conn, userID, head, info):
        super().__init__(parent)
        self.parent = parent
        self.parentform = parentform
        self.conn = conn
        self.approveclass_bool = isinstance(self, ApproveConfirmation)
        self.paymentID = info[0]
        self.userID = userID

        # Top Frame with description and user name
        self.top = tk.Frame(self, name='top_cf', padx=5, pady=5)

        # Create a frame on the canvas to contain the buttons.
        self.table_frame = tk.Frame(self.top)

        # Add info to table
        for row in zip(range(len(head)), zip(head, info)):
            if not row[1][0] == 'Краткое описание':
                self._newRow(*row)

        self._add_buttons()
        self._pack_frames()

    def _add_buttons(self):
        # Bottom Frame with buttons
        self.bottom = tk.Frame(self, name='bottom')

        bt3 = ttk.Button(self.bottom, text="Закрыть", width=10,
                         command=self.parent.destroy)
        bt3.pack(side=tk.RIGHT, padx=15, pady=5)

        if self.approveclass_bool:
            bt2 = ttk.Button(self.bottom, text="Отклонить", width=10,
                             command=lambda: self._close(False), style='ButtonRed.TButton')
            bt2.pack(side=tk.RIGHT, padx=15, pady=5)

            bt1 = ttk.Button(self.bottom, text="Утвердить", width=10,
                             command=lambda: self._close(True), style='ButtonGreen.TButton')
            bt1.pack(side=tk.RIGHT, padx=15, pady=5)


    def _pack_frames(self):
        # Pack frames
        self.top.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.table_frame.pack()
        self.pack()

    def _newRow(self, rowNumber, info):
        ''' Adds a new line to the table '''

        numberOfLines = []       # List to store number of lines needed
        columnWidths = [16, 50]  # Width of the different columns in the table - unit: text
        stringLength = []        # Lengt of the strings in the info2Add list

        #Find the length of each element in the info2Add list
        for item in info:
            stringLength.append(len(str(item)))
            numberOfLines.append(str(item).count('\n'))

        #Find the number of lines needed for each column
        for index, item in enumerate(stringLength):
            numberOfLines[index] += (ceil(item/columnWidths[index]))

        #Find the maximum number of lines needed
        lineNumber = max(numberOfLines)

        fonts = (('Calibri', 12, 'bold'), ('Calibri', 12))

        # Define labels (columns) for row
        def form_column(rowNumber, lineNumber, col_num, cell, fonts):
            col = tk.Text(self.table_frame, bg='white', padx=3)
            col.insert(1.0, cell)
            col.grid(row=rowNumber, column=col_num+1, sticky='news')
            col.configure(width=columnWidths[col_num], height=lineNumber,
                          font=fonts[col_num], state="disabled")

        for col_num, cell in enumerate(info):
            form_column(rowNumber, lineNumber, col_num, cell, fonts)


class ApproveConfirmation(DetailedRequest):
    def __init__(self, parent, parentform, conn, userID, head, info):
        super().__init__(parent, parentform, conn, userID, head, info)

    def _close(self, is_approved):
        confirmed = messagebox.askyesno(title='Подтвердите действие',
            message='{} заявку?'.format('Утвердить' if is_approved else 'Отклонить'))
        if confirmed:
            self.conn.update_confirmed(self.userID, self.paymentID, is_approved)
            self.parentform._refresh()
            self.parent.destroy()


class DetailedPreview(DetailedRequest):
    def __init__(self, parent, parentform, conn, userID, head, info):
        super().__init__(parent, parentform, conn, userID, head, info)


if __name__ == '__main__':
    from db_connect import DBConnect
    with DBConnect(server='s-kv-center-s59', db='LogisticFinance') as sql:
        try:
            app = PaymentApp(connection=sql,
                             userID=76,
                             shortusername='TestName',
                             mvz=[('mvz_a', '20511RC191'), ('mvz_b', '40900A2595')])
            app.mainloop()
        except Exception as e:
            print(e)
    input('Press Enter...')
