# -*- coding: utf-8 -*-
"""
Created on Mon Jul  8 14:00:07 2019

@author: v.shkaberda
"""
import tkinter as tk

class LabelGrid(tk.Frame):
    """
    Creates a grid of labels that have their cells populated by content.
    """
    def __init__(self, master, headers=None, content=[('Data is missing',)],
                                                      *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.headers = headers
        self.content = content
        self.content_size = (len(content), len(content[0]))
        self.im_unchecked = tk.PhotoImage(file='../resources/unchecked.png')
        self.im_checked = tk.PhotoImage(file='../resources/checked.png')
        if headers:
            self._create_headers()
        self.labels = []
        self._create_labels()

    def _create_headers(self):
        assert len(self.headers) == self.content_size[1], ('Headers and content'
                  ' should have the same number of columns')
        if isinstance(self.headers, dict):
            for j, (header, width) in enumerate(self.headers.items()):
                lb = tk.Label(self, text=header, font=('Calibri', 10, 'bold'),
                              relief='sunken', borderwidth=1, width=width)
                lb.grid(row=0, column=j, ipadx=5, sticky='nswe')
        else:
            for j, header in enumerate(self.headers):
                lb = tk.Label(self, text=header, font=('Calibri', 10, 'bold'),
                              relief='sunken', borderwidth=1)
                lb.grid(row=0, column=j, ipadx=5, sticky='nswe')

    def _create_labels(self):
        def __put_content_in_label(row, column):
            content = self.content[row][column]
            content_type = type(content).__name__
            if content_type in ('float', 'int'):
                self.labels[row][column].insert(0, content)
                if column == 0:
                    self.labels[row][column].configure(state='readonly',
                                                       justify=tk.CENTER)
            elif content_type == 'str':
                self.labels[row][column].configure(text=content, anchor=tk.W)
            elif content_type == 'bool':
                img = self.im_checked if content else self.im_unchecked
                self.labels[row][column].create_image((12, 12), image=img,
                           tag=content)
                self.labels[row][column].bind("<Button-1>", self._click)


        for i in range(self.content_size[0]):
            self.labels.append(list())
            for j in range(self.content_size[1]):
                content = self.content[i][j]
                content_type = type(content).__name__
                if content_type in ('float', 'int'):
                    self.labels[i].append(tk.Entry(self, width=10))
                elif content_type == 'str':
                    self.labels[i].append(tk.Label(self, relief='sunken',
                                                   borderwidth=1))
                elif content_type == 'bool':
                    self.labels[i].append(tk.Canvas(self, width=20, height=20))
                __put_content_in_label(i, j)
                if content_type == 'str':
                    self.labels[i][j].grid(row=i+1, column=j, sticky='we')
                else:
                    self.labels[i][j].grid(row=i+1, column=j)

    def _click(self, event):
        tags = event.widget.gettags(tk.ALL)
        if 'current' not in tags:
            return
        event.widget.delete(tk.ALL)
        if '1' in tags:
            event.widget.create_image((12, 12), image=self.im_unchecked, tag='0')
        else:
            event.widget.create_image((12, 12), image=self.im_checked, tag='1')

    def get_values(self):
        values = []
        for i in range(self.content_size[0]):
            values.append([])
            for j in range(self.content_size[1]):
                if self.labels[i][j].widgetName == 'entry':
                    values[i].append(int(self.labels[i][j].get()))
                elif self.labels[i][j].widgetName == 'canvas':
                    values[i].append(int(self.labels[i][j].gettags(tk.ALL)[0]))
        return values


if __name__ == '__main__':
    root = tk.Tk()
    label_grid = LabelGrid(root,
                           headers=('One', 'Two, but long'),
                           content=([3, True], ['my_string', False], [1, True])
                           )
    label_grid.pack()
    tk.mainloop()