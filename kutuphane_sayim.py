#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import sqlite3
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook

from tkinter import Tk, Button, Listbox, Entry, Label
from tkinter.font import Font


class Utility:
    @staticmethod
    def read_barcode(text: str):
        if len(text) > 12:
            text = text[0:12]
        return text

    @staticmethod
    def is_value_invalid(text: str):
        if text == '' or text == ' ' or len(text) < 12:
            return True
        return False

    @staticmethod
    def is_item_not_selected(selected_index: tuple):
        if selected_index is None or len(selected_index) == 0:
            return True
        return False


class DB:
    def __init__(self, path: Path):
        self.db = sqlite3.connect(path.absolute())
        self.cursor = self.db.cursor()
        self.cursor.execute('CREATE TABLE IF NOT EXISTS sayimv2 (ID INTEGER PRIMARY KEY AUTOINCREMENT, Barkod TEXT)')
        self.cursor.execute('CREATE TABLE IF NOT EXISTS logv2 (ID INTEGER PRIMARY KEY AUTOINCREMENT, Text TEXT)')
        self.db.commit()

    def close_connection(self):
        self.cursor.close()
        self.db.close()

    def read(self):
        query = 'Select * From sayimv2;'
        self.cursor.execute(query)
        list_of_items = self.cursor.fetchall()
        return list_of_items

    def get_id(self, value):
        query = 'Select ID from sayimv2 Where Barkod=?'
        self.cursor.execute(query, (value, ))
        id = self.cursor.fetchone()[0]
        return id

    def insert(self, value: str):
        query = 'Insert Into sayimv2(Barkod) Values(?)'
        self.cursor.execute(query, (value, ))
        query = "Insert Into logv2(Text) Values('{0} degeri eklendi')".format(value)
        self.cursor.execute(query)
        self.db.commit()
        id = self.get_id(value)
        return id

    def update(self, id: int, new_barkod: str):
        query = 'Update sayimv2 Set Barkod=? Where ID=?'
        self.cursor.execute(query, (new_barkod, id))
        query = "Insert Into logv2(Text) Values('{0} degeri {1} degeri ile degistirildi')".format(id, new_barkod)
        self.cursor.execute(query)
        self.db.commit()
        db_id = self.get_id(new_barkod)
        if db_id != id:
            raise Exception('Değerler birbiri ile uyuşmuyor')
        return db_id

    def delete(self, id: int):
        query = 'Delete From sayimv2 Where ID=?;'
        self.cursor.execute(query, (id, ))
        query = "Insert Into logv2(Text) Values('{0} degeri silindi')".format(id)
        self.cursor.execute(query)
        self.db.commit()


class App:
    def __init__(self, root: Tk):
        self.env_info = dict(os.environ)
        self._root = root
        self._root.title('Kütüphane Sayım Otomasyonu')
        width=482
        height=618
        screenwidth = self._root.winfo_screenwidth()
        screenheight = self._root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self._root.geometry(alignstr)
        self._root.configure(bg='white') # TODO: Will Delete Later
        self._root.resizable(width=False, height=False)

        ft = Font(family='Times',size=10)

        db_filename = 'kutuphane_sayim.db'
        if 'OS' in self.env_info and self.env_info['OS'] == 'Windows_NT':
            self.user_files_path = self.env_info['USERPROFILE']
            absolute_db_file_path = Path('{}\\{}'.format(self.user_files_path, db_filename))
            if not absolute_db_file_path.exists():
                absolute_db_file_path.touch()
            self.db = DB(absolute_db_file_path)
        else:
            self.user_files_path = self.env_info['HOME']
            absolute_db_file_path = Path('{}/{}'.format(self.user_files_path, db_filename))
            if not absolute_db_file_path.exists():
                absolute_db_file_path.touch()
            self.db = DB(absolute_db_file_path)
        init_datas = self.db.read()

        self.l_barcode=Label(self._root)
        self.l_barcode['font'] = ft
        self.l_barcode['fg'] = '#333333'
        self.l_barcode['justify'] = 'center'
        self.l_barcode['text'] = 'Barkod:'
        self.l_barcode.place(x=20,y=30,width=70,height=25)
        self.l_barcode.configure(bg='white') # TODO: Will Delete Later

        self.e_input=Entry(self._root)
        self.e_input['borderwidth'] = '1px'
        self.e_input['font'] = ft
        self.e_input['fg'] = '#333333'
        self.e_input['justify'] = 'left'
        self.e_input['text'] = ''
        self.e_input.place(x=90,y=30,width=225,height=30)
        self.e_input.bind('<Return>', self.click_insert)
        self.e_input.configure(bg='white') # TODO: Will Delete Later

        self.tb_list=Listbox(self._root)
        self.tb_list['borderwidth'] = '1px'
        self.tb_list['font'] = ft
        self.tb_list['fg'] = '#333333'
        self.tb_list['justify'] = 'left'
        self.tb_list.place(x=30,y=120,width=288,height=431)
        self.tb_list.configure(bg='white') # TODO: Will Delete Later
        if len(init_datas) > 0:
            for index, value in enumerate(init_datas):
                text = '{} | {}'.format(value[0], value[1])
                self.tb_list.insert(index, text)

        self.b_insert=Button(self._root)
        self.b_insert['bg'] = '#f0f0f0'
        self.b_insert['font'] = ft
        self.b_insert['fg'] = '#000000'
        self.b_insert['justify'] = 'center'
        self.b_insert['text'] = 'Ekle'
        self.b_insert.place(x=340,y=130,width=120,height=25)
        self.b_insert['command'] = self.click_insert

        self.b_update=Button(self._root)
        self.b_update['bg'] = '#f0f0f0'
        self.b_update['font'] = ft
        self.b_update['fg'] = '#000000'
        self.b_update['justify'] = 'center'
        self.b_update['text'] = 'Düzenle'
        self.b_update.place(x=340,y=160,width=120,height=25)
        self.b_update['command'] = self.click_update

        self.b_delete=Button(self._root)
        self.b_delete['bg'] = '#f0f0f0'
        self.b_delete['font'] = ft
        self.b_delete['fg'] = '#000000'
        self.b_delete['justify'] = 'center'
        self.b_delete['text'] = 'Sil'
        self.b_delete.place(x=340,y=190,width=120,height=25)
        self.b_delete['command'] = self.click_delete

        self.b_export_excel=Button(self._root)
        self.b_export_excel['bg'] = '#f0f0f0'
        self.b_export_excel['font'] = ft
        self.b_export_excel['fg'] = '#000000'
        self.b_export_excel['justify'] = 'center'
        self.b_export_excel['text'] = 'Excele Aktar'
        self.b_export_excel.place(x=340,y=280,width=120,height=30)
        self.b_export_excel['command'] = self.click_export_excel

        self.l_message=Label(self._root)
        self.l_message['font'] = ft
        self.l_message['fg'] = '#333333'
        self.l_message['justify'] = 'left'
        self.l_message['text'] = 'A'
        self.l_message.place(x=30,y=580,width=448,height=30)
        self.l_message.configure(bg='white') # TODO: Will Delete Later

        self.l_status=Label(self._root)
        self.l_status['font'] = ft
        self.l_status['fg'] = '#333333'
        self.l_status['justify'] = 'center'
        self.l_status['text'] = 'Durum Mesajı:'
        self.l_status.place(x=30,y=560,width=85,height=30)
        self.l_status.configure(bg='white') # TODO: Will Delete Later

    def click_insert(self, event=None):
        try:
            index = self.tb_list.size()
            value = self.e_input.get()
            if Utility.is_value_invalid(value):
                raise Exception('Lütfen geçerli bir değer giriniz.')
            value = Utility.read_barcode(value)
            db_id_value = self.db.insert(value)
            txt = '{} | {}'.format(db_id_value, value)
            self.tb_list.insert(index, txt)
            self.l_message['text'] = '{0} degeri eklendi'.format(value)
            self.e_input.delete(0, 'end')
        except Exception as ex:
            self.l_message['text'] = '{0}'.format(ex)
            self._root.bell()

    def click_update(self):
        try:
            selected_index = self.tb_list.curselection()
            if Utility.is_item_not_selected(selected_index):
                raise Exception('Lütfen listeden bir değer seçiniz')
            selected_item = self.tb_list.get(selected_index) # type(selected_item) -> tuple
            selected_db_id = int(selected_item[0])
            new_barkod = self.e_input.get()
            if Utility.is_value_invalid(new_barkod):
                raise Exception('Lütfen geçerli bir değer giriniz.')
            new_barkod = Utility.read_barcode(new_barkod)
            db_id = self.db.update(selected_db_id, new_barkod)
            txt = '{} | {}'.format(db_id, new_barkod)
            self.l_message['text'] = '{0} degeri {1} degeri ile degistirildi'.format(selected_db_id, new_barkod)
            self.tb_list.delete(selected_index)
            self.tb_list.insert(selected_index, txt)
            self.e_input.delete(0, 'end')
        except Exception as ex:
            self.l_message['text'] = '{0}'.format(ex)
            self._root.bell()

    def click_delete(self):
        try:
            selected_index = self.tb_list.curselection()
            if Utility.is_item_not_selected(selected_index):
                raise Exception('Lütfen listeden bir değer seçiniz')
            selected_item = self.tb_list.get(selected_index) # type(selected_item) -> tuple
            id = selected_item[0]
            self.db.delete(id)
            self.tb_list.delete(selected_index)
            self.l_message['text'] = '{0} degeri silindi'.format(id)
        except Exception as ex:
            self.l_message['text'] = '{0}'.format(ex)
            self._root.bell()

    def click_export_excel(self):
        try:
            size = self.tb_list.size()
            if size <= 0:
                raise Exception('En az bir öğe listede bulunmalı')
            items = self.db.read()
            dt_now = datetime.now()
            dateformat = '{}-{}-{}-{}-{}'.format(dt_now.year, dt_now.month,dt_now.day,dt_now.hour,dt_now.minute)

            my_wb = Workbook()
            my_sheet = my_wb.active
            for index, item in enumerate(items):
                new_column = my_sheet.cell(row = index+1, column = 1)
                new_column.value = item[1]
            if 'OS' in self.env_info and self.env_info['OS'] == 'Windows_NT':
                filename = '{}\\sayim_cikti_{}.xlsx'.format(self.user_files_path, dateformat)
            else:
                filename = '{}/sayim_cikti_{}.xlsx'.format(self.user_files_path, dateformat)
            my_wb.save(filename)
            self.l_message['text'] = 'Dosya Adı: {}'.format(size, filename)
        except Exception as ex:
            self.l_message['text'] = '{}'.format(ex)
            self._root.bell()


if __name__ == '__main__':
    root = Tk()
    app = App(root)
    root.mainloop()

