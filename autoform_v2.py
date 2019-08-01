from tkinter import *
from tkinter import ttk
from tkinter import messagebox as ms
from openpyxl import *
from time import strftime as time
import os
import sys


# API
def total(name, index, mode):
    vlr_total_01.set(qntd_01.get() * vlr_un_01.get())
    vlr_total_02.set(qntd_02.get() * vlr_un_02.get())
    vlr_total_03.set(qntd_03.get() * vlr_un_03.get())
    vlr_total_04.set(qntd_04.get() * vlr_un_04.get())
    vlr_total_05.set(qntd_05.get() * vlr_un_05.get())
    vlr_total_06.set(qntd_06.get() * vlr_un_06.get())
    vlr_total_07.set(qntd_07.get() * vlr_un_07.get())
    vlr_total_08.set(qntd_08.get() * vlr_un_08.get())
    vlr_total_09.set(qntd_09.get() * vlr_un_09.get())
    vlr_total_10.set(qntd_10.get() * vlr_un_10.get())
    vlr_total_11.set(qntd_11.get() * vlr_un_11.get())
    vlr_total_12.set(qntd_12.get() * vlr_un_12.get())
    vlr_total_13.set(qntd_13.get() * vlr_un_13.get())
    vlr_total_14.set(qntd_14.get() * vlr_un_14.get())
    vlr_total_15.set(qntd_15.get() * vlr_un_15.get())
    vlr_total_16.set(qntd_16.get() * vlr_un_16.get())
    vlr_total_17.set(qntd_17.get() * vlr_un_17.get())
    vlr_total_18.set(qntd_18.get() * vlr_un_18.get())
    vlr_total_19.set(qntd_19.get() * vlr_un_19.get())
    vlr_total_20.set(qntd_20.get() * vlr_un_20.get())
    vlr_total_21.set(qntd_01.get() * vlr_un_21.get())
    vlr_total_22.set(qntd_02.get() * vlr_un_22.get())
    vlr_total_23.set(qntd_03.get() * vlr_un_23.get())
    vlr_total_24.set(qntd_04.get() * vlr_un_24.get())
    vlr_total_25.set(qntd_05.get() * vlr_un_25.get())
    vlr_total_26.set(qntd_06.get() * vlr_un_26.get())
    vlr_total_27.set(qntd_07.get() * vlr_un_27.get())
    vlr_total_28.set(qntd_08.get() * vlr_un_28.get())
    vlr_total_29.set(qntd_09.get() * vlr_un_29.get())
    vlr_total_30.set(qntd_10.get() * vlr_un_30.get())
    vlr_total.set(vlr_total_01.get() +
                  vlr_total_02.get() +
                  vlr_total_03.get() +
                  vlr_total_04.get() +
                  vlr_total_05.get() +
                  vlr_total_06.get() +
                  vlr_total_07.get() +
                  vlr_total_08.get() +
                  vlr_total_09.get() +
                  vlr_total_10.get() +
                  vlr_total_11.get() +
                  vlr_total_12.get() +
                  vlr_total_13.get() +
                  vlr_total_14.get() +
                  vlr_total_15.get() +
                  vlr_total_16.get() +
                  vlr_total_17.get() +
                  vlr_total_18.get() +
                  vlr_total_19.get() +
                  vlr_total_20.get() +
                  vlr_total_21.get() +
                  vlr_total_22.get() +
                  vlr_total_23.get() +
                  vlr_total_24.get() +
                  vlr_total_25.get() +
                  vlr_total_26.get() +
                  vlr_total_27.get() +
                  vlr_total_28.get() +
                  vlr_total_29.get() +
                  vlr_total_30.get())
    vlr_total_final.set((vlr_total.get()) - (vlr_total.get() * (desconto.get() / 100)))


def refresh_clock():
    global time_1
    time_2 = time('%d-%m-%Y - %Hh%Mm%Ss')
    if time_2 != time_1:
        time_1 = time_2

    clock.config(text=time_1)
    clock.after(200, refresh_clock)


def file_save():
    os.chdir(save_dir)
    # CADASTRO
    sheet['E4'] = time('%d-%m-%Y - %Hh%Mm%Ss')
    sheet['E5'] = entry_name.get().upper()
    sheet['E6'] = entry_address.get().upper()
    sheet['E7'] = entry_phone.get().upper()
    sheet['E8'] = entry_email.get()
    # PRODUTOS
    sheet['A11'] = prod_01.get().upper()
    sheet['A12'] = prod_02.get().upper()
    sheet['A13'] = prod_03.get().upper()
    sheet['A14'] = prod_04.get().upper()
    sheet['A15'] = prod_05.get().upper()
    sheet['A16'] = prod_06.get().upper()
    sheet['A17'] = prod_07.get().upper()
    sheet['A18'] = prod_08.get().upper()
    sheet['A19'] = prod_09.get().upper()
    sheet['A20'] = prod_10.get().upper()
    sheet['A21'] = prod_11.get().upper()
    sheet['A22'] = prod_12.get().upper()
    sheet['A23'] = prod_13.get().upper()
    sheet['A24'] = prod_14.get().upper()
    sheet['A25'] = prod_15.get().upper()
    sheet['A26'] = prod_16.get().upper()
    sheet['A27'] = prod_17.get().upper()
    sheet['A28'] = prod_18.get().upper()
    sheet['A29'] = prod_19.get().upper()
    sheet['A30'] = prod_20.get().upper()
    sheet['A31'] = prod_21.get().upper()
    sheet['A32'] = prod_22.get().upper()
    sheet['A33'] = prod_23.get().upper()
    sheet['A34'] = prod_24.get().upper()
    sheet['A35'] = prod_25.get().upper()
    sheet['A36'] = prod_26.get().upper()
    sheet['A37'] = prod_27.get().upper()
    sheet['A38'] = prod_28.get().upper()
    sheet['A39'] = prod_29.get().upper()
    sheet['A40'] = prod_30.get().upper()
    # QUANTIDADE
    sheet['E11'] = qntd_01.get()
    sheet['E12'] = qntd_02.get()
    sheet['E13'] = qntd_03.get()
    sheet['E14'] = qntd_04.get()
    sheet['E15'] = qntd_05.get()
    sheet['E16'] = qntd_06.get()
    sheet['E17'] = qntd_07.get()
    sheet['E18'] = qntd_08.get()
    sheet['E19'] = qntd_09.get()
    sheet['E20'] = qntd_10.get()
    sheet['E21'] = qntd_11.get()
    sheet['E22'] = qntd_12.get()
    sheet['E23'] = qntd_13.get()
    sheet['E24'] = qntd_14.get()
    sheet['E25'] = qntd_15.get()
    sheet['E26'] = qntd_16.get()
    sheet['E27'] = qntd_17.get()
    sheet['E28'] = qntd_18.get()
    sheet['E29'] = qntd_19.get()
    sheet['E30'] = qntd_20.get()
    sheet['E31'] = qntd_21.get()
    sheet['E32'] = qntd_22.get()
    sheet['E33'] = qntd_23.get()
    sheet['E34'] = qntd_24.get()
    sheet['E35'] = qntd_25.get()
    sheet['E36'] = qntd_26.get()
    sheet['E37'] = qntd_27.get()
    sheet['E38'] = qntd_28.get()
    sheet['E39'] = qntd_29.get()
    sheet['E40'] = qntd_30.get()
    # VALOR UNID.
    sheet['G11'] = vlr_un_01.get()
    sheet['G12'] = vlr_un_02.get()
    sheet['G13'] = vlr_un_03.get()
    sheet['G14'] = vlr_un_04.get()
    sheet['G15'] = vlr_un_05.get()
    sheet['G16'] = vlr_un_06.get()
    sheet['G17'] = vlr_un_07.get()
    sheet['G18'] = vlr_un_08.get()
    sheet['G19'] = vlr_un_09.get()
    sheet['G20'] = vlr_un_10.get()
    sheet['G21'] = vlr_un_11.get()
    sheet['G22'] = vlr_un_12.get()
    sheet['G23'] = vlr_un_13.get()
    sheet['G24'] = vlr_un_14.get()
    sheet['G25'] = vlr_un_15.get()
    sheet['G26'] = vlr_un_16.get()
    sheet['G27'] = vlr_un_17.get()
    sheet['G28'] = vlr_un_18.get()
    sheet['G29'] = vlr_un_19.get()
    sheet['G30'] = vlr_un_20.get()
    sheet['G31'] = vlr_un_21.get()
    sheet['G32'] = vlr_un_22.get()
    sheet['G33'] = vlr_un_23.get()
    sheet['G34'] = vlr_un_24.get()
    sheet['G35'] = vlr_un_25.get()
    sheet['G36'] = vlr_un_26.get()
    sheet['G37'] = vlr_un_27.get()
    sheet['G38'] = vlr_un_28.get()
    sheet['G39'] = vlr_un_29.get()
    sheet['G40'] = vlr_un_30.get()
    # PAGAMENTO
    sheet['H42'] = forma_pagamento.get()
    sheet['H43'] = desconto.get()
    sheet['H44'] = vlr_total_final.get()

    file_name = str(time('%d-%m-%Y - %Hh%Mm%Ss') + '- Nota.xlsx')
    wb.save(file_name)

    ms.showinfo('SALVAR', 'A Nota foi salva com sucesso!')

    clear()


def file_quit():
    if ms.askyesno('SAIR', 'Você gostaria de encerrar o programa?'):
        sys.exit()


def clear():
    vlr_total.trace('w', total)
    vlr_total.set(0)

    vlr_total_final.trace('w', total)
    vlr_total_final.set(0)

    desconto.trace('w', total)
    desconto.set(0)

    qntd_01.trace('w', total)
    qntd_02.trace('w', total)
    qntd_03.trace('w', total)
    qntd_04.trace('w', total)
    qntd_05.trace('w', total)
    qntd_06.trace('w', total)
    qntd_07.trace('w', total)
    qntd_08.trace('w', total)
    qntd_09.trace('w', total)
    qntd_10.trace('w', total)
    qntd_11.trace('w', total)
    qntd_12.trace('w', total)
    qntd_13.trace('w', total)
    qntd_14.trace('w', total)
    qntd_15.trace('w', total)
    qntd_16.trace('w', total)
    qntd_17.trace('w', total)
    qntd_18.trace('w', total)
    qntd_19.trace('w', total)
    qntd_20.trace('w', total)
    qntd_21.trace('w', total)
    qntd_22.trace('w', total)
    qntd_23.trace('w', total)
    qntd_24.trace('w', total)
    qntd_25.trace('w', total)
    qntd_26.trace('w', total)
    qntd_27.trace('w', total)
    qntd_28.trace('w', total)
    qntd_29.trace('w', total)
    qntd_30.trace('w', total)

    vlr_un_01.trace('w', total)
    vlr_un_02.trace('w', total)
    vlr_un_03.trace('w', total)
    vlr_un_04.trace('w', total)
    vlr_un_05.trace('w', total)
    vlr_un_06.trace('w', total)
    vlr_un_07.trace('w', total)
    vlr_un_08.trace('w', total)
    vlr_un_09.trace('w', total)
    vlr_un_10.trace('w', total)
    vlr_un_11.trace('w', total)
    vlr_un_12.trace('w', total)
    vlr_un_13.trace('w', total)
    vlr_un_14.trace('w', total)
    vlr_un_15.trace('w', total)
    vlr_un_16.trace('w', total)
    vlr_un_17.trace('w', total)
    vlr_un_18.trace('w', total)
    vlr_un_19.trace('w', total)
    vlr_un_20.trace('w', total)
    vlr_un_21.trace('w', total)
    vlr_un_22.trace('w', total)
    vlr_un_23.trace('w', total)
    vlr_un_24.trace('w', total)
    vlr_un_25.trace('w', total)
    vlr_un_26.trace('w', total)
    vlr_un_27.trace('w', total)
    vlr_un_28.trace('w', total)
    vlr_un_29.trace('w', total)
    vlr_un_30.trace('w', total)

    qntd_01.set(0)
    qntd_02.set(0)
    qntd_03.set(0)
    qntd_04.set(0)
    qntd_05.set(0)
    qntd_06.set(0)
    qntd_07.set(0)
    qntd_08.set(0)
    qntd_09.set(0)
    qntd_10.set(0)
    qntd_11.set(0)
    qntd_12.set(0)
    qntd_13.set(0)
    qntd_14.set(0)
    qntd_15.set(0)
    qntd_16.set(0)
    qntd_17.set(0)
    qntd_18.set(0)
    qntd_19.set(0)
    qntd_20.set(0)
    qntd_21.set(0)
    qntd_22.set(0)
    qntd_23.set(0)
    qntd_24.set(0)
    qntd_25.set(0)
    qntd_26.set(0)
    qntd_27.set(0)
    qntd_28.set(0)
    qntd_29.set(0)
    qntd_30.set(0)

    vlr_un_01.set(0)
    vlr_un_02.set(0)
    vlr_un_03.set(0)
    vlr_un_04.set(0)
    vlr_un_05.set(0)
    vlr_un_06.set(0)
    vlr_un_07.set(0)
    vlr_un_08.set(0)
    vlr_un_09.set(0)
    vlr_un_10.set(0)
    vlr_un_11.set(0)
    vlr_un_12.set(0)
    vlr_un_13.set(0)
    vlr_un_14.set(0)
    vlr_un_15.set(0)
    vlr_un_16.set(0)
    vlr_un_17.set(0)
    vlr_un_18.set(0)
    vlr_un_19.set(0)
    vlr_un_20.set(0)
    vlr_un_21.set(0)
    vlr_un_22.set(0)
    vlr_un_23.set(0)
    vlr_un_24.set(0)
    vlr_un_25.set(0)
    vlr_un_26.set(0)
    vlr_un_27.set(0)
    vlr_un_28.set(0)
    vlr_un_29.set(0)
    vlr_un_30.set(0)

    prod_01.set('')
    prod_02.set('')
    prod_03.set('')
    prod_04.set('')
    prod_05.set('')
    prod_06.set('')
    prod_07.set('')
    prod_08.set('')
    prod_09.set('')
    prod_10.set('')
    prod_11.set('')
    prod_12.set('')
    prod_13.set('')
    prod_14.set('')
    prod_15.set('')
    prod_16.set('')
    prod_17.set('')
    prod_18.set('')
    prod_19.set('')
    prod_20.set('')
    prod_21.set('')
    prod_22.set('')
    prod_23.set('')
    prod_24.set('')
    prod_25.set('')
    prod_26.set('')
    prod_27.set('')
    prod_28.set('')
    prod_29.set('')
    prod_30.set('')

    entry_name.set('')
    entry_address.set('')
    entry_phone.set('')
    entry_email.set('')

    forma_pagamento.set('')


# INIT
template = (os.path.join(os.path.join(os.getenv('USERPROFILE')), r'Desktop\autoform\bin'))
save_dir = (os.path.join(os.path.join(os.getenv('USERPROFILE')), r'Desktop\autoform\saved'))

os.chdir(template)

wb = load_workbook('template.xlsx')
sheet = wb['Plan1']

root = Tk()

time_1 = ''
clock = Label(root, font=('Verdana', 10, 'bold'), borderwidth=2, relief='groove', width=27)

# MENU
menu_bar = Menu(root)
menu_bar.add_command(label='LIMPAR NOTA', command=clear)
menu_bar.add_command(label='SALVAR', command=file_save)
menu_bar.add_command(label='SAIR', command=file_quit)


# REGISTER
Label(root, text='Data da Geração do Documento :', borderwidth=2, relief='groove', width=30).grid(row=0, column=0)
clock.grid(row=0, column=1, sticky=W)
Label(root, text='Nome do Cliente :', borderwidth=2, relief='groove', width=30).grid(row=1, column=0)
entry_name = StringVar()
Entry(root, textvariable=entry_name, width=40).grid(row=1, column=1, sticky=W)
Label(root, text='Endereço :', borderwidth=2, relief='groove', width=30).grid(row=2, column=0)
entry_address = StringVar()
Entry(root, textvariable=entry_address, width=40).grid(row=2, column=1, sticky=W)
Label(root, text='Telefone :', borderwidth=2, relief='groove', width=30).grid(row=3, column=0)
entry_phone = StringVar()
Entry(root, textvariable=entry_phone, width=40).grid(row=3, column=1, sticky=W)
Label(root, text='E-Mail :', borderwidth=2, relief='groove', width=30).grid(row=4, column=0)
entry_email = StringVar()
Entry(root, textvariable=entry_email, width=40).grid(row=4, column=1, sticky=W)

# PRODUTOS
Label(root, text='PRODUTOS :', borderwidth=2, relief='groove', width=40).grid(row=5, column=1)
Label(root, text='QUANTIDADE :', borderwidth=2, relief='groove', width=15).grid(row=5, column=2)
Label(root, text='VALOR UNID.:', borderwidth=2, relief='groove', width=15).grid(row=5, column=3)
Label(root, text='VALOR TOTAL :', borderwidth=2, relief='groove', width=15).grid(row=5, column=4)

# 1
prod_01 = StringVar()
qntd_01 = IntVar()
vlr_un_01 = DoubleVar()
vlr_total_01 = DoubleVar()

Entry(root, textvariable=prod_01, width=40).grid(row=6, column=1)
Entry(root, textvariable=qntd_01, width=10).grid(row=6, column=2)
Entry(root, textvariable=vlr_un_01, width=10).grid(row=6, column=3)
Label(root, textvariable=vlr_total_01, borderwidth=3, relief='raised', width=15).grid(row=6, column=4)

# 2
prod_02 = StringVar()
qntd_02 = IntVar()
vlr_un_02 = DoubleVar()
vlr_total_02 = DoubleVar()

Entry(root, textvariable=prod_02, width=40).grid(row=7, column=1)
Entry(root, textvariable=qntd_02, width=10).grid(row=7, column=2)
Entry(root, textvariable=vlr_un_02, width=10).grid(row=7, column=3)
Label(root, textvariable=vlr_total_02, borderwidth=3, relief='raised', width=15).grid(row=7, column=4)

# 3
prod_03 = StringVar()
qntd_03 = IntVar()
vlr_un_03 = DoubleVar()
vlr_total_03 = DoubleVar()

Entry(root, textvariable=prod_03, width=40).grid(row=8, column=1)
Entry(root, textvariable=qntd_03, width=10).grid(row=8, column=2)
Entry(root, textvariable=vlr_un_03, width=10).grid(row=8, column=3)
Label(root, textvariable=vlr_total_03, borderwidth=3, relief='raised', width=15).grid(row=8, column=4)

# 4
prod_04 = StringVar()
qntd_04 = IntVar()
vlr_un_04 = DoubleVar()
vlr_total_04 = DoubleVar()

Entry(root, textvariable=prod_04, width=40).grid(row=9, column=1)
Entry(root, textvariable=qntd_04, width=10).grid(row=9, column=2)
Entry(root, textvariable=vlr_un_04, width=10).grid(row=9, column=3)
Label(root, textvariable=vlr_total_04, borderwidth=3, relief='raised', width=15).grid(row=9, column=4)

# 5
prod_05 = StringVar()
qntd_05 = IntVar()
vlr_un_05 = DoubleVar()
vlr_total_05 = DoubleVar()

Entry(root, textvariable=prod_05, width=40).grid(row=10, column=1)
Entry(root, textvariable=qntd_05, width=10).grid(row=10, column=2)
Entry(root, textvariable=vlr_un_05, width=10).grid(row=10, column=3)
Label(root, textvariable=vlr_total_05, borderwidth=3, relief='raised', width=15).grid(row=10, column=4)

# 6
prod_06 = StringVar()
qntd_06 = IntVar()
vlr_un_06 = DoubleVar()
vlr_total_06 = DoubleVar()

Entry(root, textvariable=prod_06, width=40).grid(row=11, column=1)
Entry(root, textvariable=qntd_06, width=10).grid(row=11, column=2)
Entry(root, textvariable=vlr_un_06, width=10).grid(row=11, column=3)
Label(root, textvariable=vlr_total_06, borderwidth=3, relief='raised', width=15).grid(row=11, column=4)

# 7
prod_07 = StringVar()
qntd_07 = IntVar()
vlr_un_07 = DoubleVar()
vlr_total_07 = DoubleVar()

Entry(root, textvariable=prod_07, width=40).grid(row=12, column=1)
Entry(root, textvariable=qntd_07, width=10).grid(row=12, column=2)
Entry(root, textvariable=vlr_un_07, width=10).grid(row=12, column=3)
Label(root, textvariable=vlr_total_07, borderwidth=3, relief='raised', width=15).grid(row=12, column=4)

# 8
prod_08 = StringVar()
qntd_08 = IntVar()
vlr_un_08 = DoubleVar()
vlr_total_08 = DoubleVar()

Entry(root, textvariable=prod_08, width=40).grid(row=13, column=1)
Entry(root, textvariable=qntd_08, width=10).grid(row=13, column=2)
Entry(root, textvariable=vlr_un_08, width=10).grid(row=13, column=3)
Label(root, textvariable=vlr_total_08, borderwidth=3, relief='raised', width=15).grid(row=13, column=4)

# 9
prod_09 = StringVar()
qntd_09 = IntVar()
vlr_un_09 = DoubleVar()
vlr_total_09 = DoubleVar()

Entry(root, textvariable=prod_09, width=40).grid(row=14, column=1)
Entry(root, textvariable=qntd_09, width=10).grid(row=14, column=2)
Entry(root, textvariable=vlr_un_09, width=10).grid(row=14, column=3)
Label(root, textvariable=vlr_total_09, borderwidth=3, relief='raised', width=15).grid(row=14, column=4)

# 10
prod_10 = StringVar()
qntd_10 = IntVar()
vlr_un_10 = DoubleVar()
vlr_total_10 = DoubleVar()

Entry(root, textvariable=prod_10, width=40).grid(row=15, column=1)
Entry(root, textvariable=qntd_10, width=10).grid(row=15, column=2)
Entry(root, textvariable=vlr_un_10, width=10).grid(row=15, column=3)
Label(root, textvariable=vlr_total_10, borderwidth=3, relief='raised', width=15).grid(row=15, column=4)

# 11
prod_11 = StringVar()
qntd_11 = IntVar()
vlr_un_11 = DoubleVar()
vlr_total_11 = DoubleVar()

Entry(root, textvariable=prod_11, width=40).grid(row=16, column=1)
Entry(root, textvariable=qntd_11, width=10).grid(row=16, column=2)
Entry(root, textvariable=vlr_un_11, width=10).grid(row=16, column=3)
Label(root, textvariable=vlr_total_11, borderwidth=3, relief='raised', width=15).grid(row=16, column=4)

# 12
prod_12 = StringVar()
qntd_12 = IntVar()
vlr_un_12 = DoubleVar()
vlr_total_12 = DoubleVar()

Entry(root, textvariable=prod_12, width=40).grid(row=17, column=1)
Entry(root, textvariable=qntd_12, width=10).grid(row=17, column=2)
Entry(root, textvariable=vlr_un_12, width=10).grid(row=17, column=3)
Label(root, textvariable=vlr_total_12, borderwidth=3, relief='raised', width=15).grid(row=17, column=4)

# 13
prod_13 = StringVar()
qntd_13 = IntVar()
vlr_un_13 = DoubleVar()
vlr_total_13 = DoubleVar()

Entry(root, textvariable=prod_13, width=40).grid(row=18, column=1)
Entry(root, textvariable=qntd_13, width=10).grid(row=18, column=2)
Entry(root, textvariable=vlr_un_13, width=10).grid(row=18, column=3)
Label(root, textvariable=vlr_total_13, borderwidth=3, relief='raised', width=15).grid(row=18, column=4)

# 14
prod_14 = StringVar()
qntd_14 = IntVar()
vlr_un_14 = DoubleVar()
vlr_total_14 = DoubleVar()

Entry(root, textvariable=prod_14, width=40).grid(row=19, column=1)
Entry(root, textvariable=qntd_14, width=10).grid(row=19, column=2)
Entry(root, textvariable=vlr_un_14, width=10).grid(row=19, column=3)
Label(root, textvariable=vlr_total_14, borderwidth=3, relief='raised', width=15).grid(row=19, column=4)

# 15
prod_15 = StringVar()
qntd_15 = IntVar()
vlr_un_15 = DoubleVar()
vlr_total_15 = DoubleVar()

Entry(root, textvariable=prod_15, width=40).grid(row=20, column=1)
Entry(root, textvariable=qntd_15, width=10).grid(row=20, column=2)
Entry(root, textvariable=vlr_un_15, width=10).grid(row=20, column=3)
Label(root, textvariable=vlr_total_15, borderwidth=3, relief='raised', width=15).grid(row=20, column=4)

# 16
prod_16 = StringVar()
qntd_16 = IntVar()
vlr_un_16 = DoubleVar()
vlr_total_16 = DoubleVar()

Entry(root, textvariable=prod_16, width=40).grid(row=21, column=1)
Entry(root, textvariable=qntd_16, width=10).grid(row=21, column=2)
Entry(root, textvariable=vlr_un_16, width=10).grid(row=21, column=3)
Label(root, textvariable=vlr_total_16, borderwidth=3, relief='raised', width=15).grid(row=21, column=4)

# 17
prod_17 = StringVar()
qntd_17 = IntVar()
vlr_un_17 = DoubleVar()
vlr_total_17 = DoubleVar()

Entry(root, textvariable=prod_17, width=40).grid(row=22, column=1)
Entry(root, textvariable=qntd_17, width=10).grid(row=22, column=2)
Entry(root, textvariable=vlr_un_17, width=10).grid(row=22, column=3)
Label(root, textvariable=vlr_total_17, borderwidth=3, relief='raised', width=15).grid(row=22, column=4)

# 18
prod_18 = StringVar()
qntd_18 = IntVar()
vlr_un_18 = DoubleVar()
vlr_total_18 = DoubleVar()

Entry(root, textvariable=prod_18, width=40).grid(row=23, column=1)
Entry(root, textvariable=qntd_18, width=10).grid(row=23, column=2)
Entry(root, textvariable=vlr_un_18, width=10).grid(row=23, column=3)
Label(root, textvariable=vlr_total_18, borderwidth=3, relief='raised', width=15).grid(row=23, column=4)

# 19
prod_19 = StringVar()
qntd_19 = IntVar()
vlr_un_19 = DoubleVar()
vlr_total_19 = DoubleVar()

Entry(root, textvariable=prod_19, width=40).grid(row=24, column=1)
Entry(root, textvariable=qntd_19, width=10).grid(row=24, column=2)
Entry(root, textvariable=vlr_un_19, width=10).grid(row=24, column=3)
Label(root, textvariable=vlr_total_19, borderwidth=3, relief='raised', width=15).grid(row=24, column=4)

# 20
prod_20 = StringVar()
qntd_20 = IntVar()
vlr_un_20 = DoubleVar()
vlr_total_20 = DoubleVar()

Entry(root, textvariable=prod_20, width=40).grid(row=25, column=1)
Entry(root, textvariable=qntd_20, width=10).grid(row=25, column=2)
Entry(root, textvariable=vlr_un_20, width=10).grid(row=25, column=3)
Label(root, textvariable=vlr_total_20, borderwidth=3, relief='raised', width=15).grid(row=25, column=4)

# 21
prod_21 = StringVar()
qntd_21 = IntVar()
vlr_un_21 = DoubleVar()
vlr_total_21 = DoubleVar()

Entry(root, textvariable=prod_21, width=40).grid(row=26, column=1)
Entry(root, textvariable=qntd_21, width=10).grid(row=26, column=2)
Entry(root, textvariable=vlr_un_21, width=10).grid(row=26, column=3)
Label(root, textvariable=vlr_total_21, borderwidth=3, relief='raised', width=15).grid(row=26, column=4)

# 22
prod_22 = StringVar()
qntd_22 = IntVar()
vlr_un_22 = DoubleVar()
vlr_total_22 = DoubleVar()

Entry(root, textvariable=prod_22, width=40).grid(row=27, column=1)
Entry(root, textvariable=qntd_22, width=10).grid(row=27, column=2)
Entry(root, textvariable=vlr_un_22, width=10).grid(row=27, column=3)
Label(root, textvariable=vlr_total_22, borderwidth=3, relief='raised', width=15).grid(row=27, column=4)

# 23
prod_23 = StringVar()
qntd_23 = IntVar()
vlr_un_23 = DoubleVar()
vlr_total_23 = DoubleVar()

Entry(root, textvariable=prod_23, width=40).grid(row=28, column=1)
Entry(root, textvariable=qntd_23, width=10).grid(row=28, column=2)
Entry(root, textvariable=vlr_un_23, width=10).grid(row=28, column=3)
Label(root, textvariable=vlr_total_23, borderwidth=3, relief='raised', width=15).grid(row=28, column=4)

# 24
prod_24 = StringVar()
qntd_24 = IntVar()
vlr_un_24 = DoubleVar()
vlr_total_24 = DoubleVar()

Entry(root, textvariable=prod_24, width=40).grid(row=29, column=1)
Entry(root, textvariable=qntd_24, width=10).grid(row=29, column=2)
Entry(root, textvariable=vlr_un_24, width=10).grid(row=29, column=3)
Label(root, textvariable=vlr_total_24, borderwidth=3, relief='raised', width=15).grid(row=29, column=4)

# 25
prod_25 = StringVar()
qntd_25 = IntVar()
vlr_un_25 = DoubleVar()
vlr_total_25 = DoubleVar()

Entry(root, textvariable=prod_25, width=40).grid(row=30, column=1)
Entry(root, textvariable=qntd_25, width=10).grid(row=30, column=2)
Entry(root, textvariable=vlr_un_25, width=10).grid(row=30, column=3)
Label(root, textvariable=vlr_total_25, borderwidth=3, relief='raised', width=15).grid(row=30, column=4)

# 26
prod_26 = StringVar()
qntd_26 = IntVar()
vlr_un_26 = DoubleVar()
vlr_total_26 = DoubleVar()

Entry(root, textvariable=prod_26, width=40).grid(row=31, column=1)
Entry(root, textvariable=qntd_26, width=10).grid(row=31, column=2)
Entry(root, textvariable=vlr_un_26, width=10).grid(row=31, column=3)
Label(root, textvariable=vlr_total_26, borderwidth=3, relief='raised', width=15).grid(row=31, column=4)

# 27
prod_27 = StringVar()
qntd_27 = IntVar()
vlr_un_27 = DoubleVar()
vlr_total_27 = DoubleVar()

Entry(root, textvariable=prod_27, width=40).grid(row=32, column=1)
Entry(root, textvariable=qntd_27, width=10).grid(row=32, column=2)
Entry(root, textvariable=vlr_un_27, width=10).grid(row=32, column=3)
Label(root, textvariable=vlr_total_27, borderwidth=3, relief='raised', width=15).grid(row=32, column=4)

# 28
prod_28 = StringVar()
qntd_28 = IntVar()
vlr_un_28 = DoubleVar()
vlr_total_28 = DoubleVar()

Entry(root, textvariable=prod_28, width=40).grid(row=33, column=1)
Entry(root, textvariable=qntd_28, width=10).grid(row=33, column=2)
Entry(root, textvariable=vlr_un_28, width=10).grid(row=33, column=3)
Label(root, textvariable=vlr_total_28, borderwidth=3, relief='raised', width=15).grid(row=33, column=4)

# 29
prod_29 = StringVar()
qntd_29 = IntVar()
vlr_un_29 = DoubleVar()
vlr_total_29 = DoubleVar()

Entry(root, textvariable=prod_29, width=40).grid(row=34, column=1)
Entry(root, textvariable=qntd_29, width=10).grid(row=34, column=2)
Entry(root, textvariable=vlr_un_29, width=10).grid(row=34, column=3)
Label(root, textvariable=vlr_total_29, borderwidth=3, relief='raised', width=15).grid(row=34, column=4)

# 30
prod_30 = StringVar()
qntd_30 = IntVar()
vlr_un_30 = DoubleVar()
vlr_total_30 = DoubleVar()

Entry(root, textvariable=prod_30, width=40).grid(row=35, column=1)
Entry(root, textvariable=qntd_30, width=10).grid(row=35, column=2)
Entry(root, textvariable=vlr_un_30, width=10).grid(row=35, column=3)
Label(root, textvariable=vlr_total_30, borderwidth=3, relief='raised', width=15).grid(row=35, column=4)

# TOTAL DA COMPRA
vlr_total = DoubleVar()
Label(root, text='VALOR TOTAL DA COMPRA :', borderwidth=2, relief='groove', width=25).grid(row=5, column=5)
Label(root, textvariable=vlr_total, borderwidth=3, relief='raised', width=15).grid(row=6, column=5)

# FORMA DE PAGAMENTO
Label(root, text='FORMA DE PAGAMENTO :', borderwidth=2, relief='groove', width=25).grid(row=7, column=5)
pagamento_list = ['À VISTA', 'CHEQUE', 'CARTÃO 2x', 'CARTÃO 3x', 'CARTÃO 4x', 'CARTÃO 5x', 'CARTÃO 6x']
forma_pagamento = ttk.Combobox(root, values=pagamento_list)
forma_pagamento.grid(row=8, column=5)
Label(root, text='DESCONTO (%):', borderwidth=2, relief='groove', width=25).grid(row=9, column=5)
desconto = DoubleVar()
Entry(root, textvariable=desconto).grid(row=10, column=5)

vlr_total_final = DoubleVar()
Label(root, text='VALOR À PAGAR :', borderwidth=2, relief='groove', width=25).grid(row=11, column=5)
Label(root, textvariable=vlr_total_final, borderwidth=3, relief='raised', width=15).grid(row=12, column=5)



# CONFIG
if __name__ == '__main__':
    clear()
    refresh_clock()
    root.title('AUTOFORM V2.0 - The Pythonic Boogaloo')
    root.attributes('-fullscreen', True)
    root.config(menu=menu_bar)
    root.mainloop()
