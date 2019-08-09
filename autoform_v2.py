from tkinter import *
from tkinter import ttk
from tkinter import messagebox as ms
from openpyxl import *
from time import strftime as time
import os
import sys


# API
    
def addProdutos(inputList,number = 1):
    """This functions will a product in the inputList.
        The default number of products to be added is 1,
        but it can be used to add multiple products at once,
        setting the 'number' variable to the number of products you want 
    """
    index = len(inputList) # Used to know what is the number of the product.
    for i in range(number):
        #! I am creating an array that contains every value to be stored,
        #! keeping the same order the original developer wanted.
        tmpData =[StringVar(),IntVar(),DoubleVar(),DoubleVar()]
        tmpFields = [Entry(root, textvariable=tmpData[0], width=40),Entry(root, textvariable=tmpData[1], width=10),Entry(root, textvariable=tmpData[2], width=10),Label(root, textvariable=tmpData[3], borderwidth=3, relief='raised', width=15)]
        for el in enumerate(tmpFields,start=1):
            el[1].grid(row=index + 6, column=el[0])
            el[1].bind('<KeyRelease>', total)
        tmpProd = [tmpData,tmpFields]
        inputList.append(tmpProd)
        index += 1 #adding 1, so we don't have to read the len(inputList) everytime

def total(event=None,*args):
    tmpTotal = 0
    for x in listProdutos:
        x[0][3].set(x[0][1].get() * x[0][2].get())
        tmpTotal += x[0][3].get()
    vlr_total.set(tmpTotal)
    vlr_total_final.set((vlr_total.get()) - (vlr_total.get() * (desconto.get() / 100)))


def refresh_clock():
    global time_1
    time_2 = time('%d-%m-%Y - %Hh%Mm%Ss')
    if time_2 != time_1:
        time_1 = time_2

    clock.config(text=time_1)
    clock.after(200, refresh_clock)


def file_save():
    #! I will not change this, because i will not work with excel now
    #! BUT, there is a range method that could be used, if th number of products is defined
    
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
    #! setting every value to '' ... this should clean the fields

    for x in listProdutos:
        for y in x[0]:
            y.set('')
    
    vlr_total.trace('w', total)
    vlr_total.set(0)

    vlr_total_final.trace('w', total)
    vlr_total_final.set(0)

    desconto.trace('w', total)
    desconto.set(0)

    entry_name.set('')
    entry_address.set('')
    entry_phone.set('')
    entry_email.set('')

    forma_pagamento.set('')


# INIT
#! Changed the path, to get the files from relative path.
#! In this way, the program should work in every directory you put it. 
template = (os.path.dirname(os.path.realpath(__file__)) + r'\autoform\bin')
save_dir = (os.path.dirname(os.path.realpath(__file__)) + r'\autoform\saved')

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

#! Adding the list of products here
listProdutos = []
addProdutos(listProdutos,30)

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
