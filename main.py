# File: main.py
import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication,QFileDialog,QMessageBox
from PySide6.QtCore import QFile, QIODevice
import os
from openpyxl import load_workbook
import xlrd
import re



#some init vars
path_file1 = ''
path_file2 = ''
path_file3 = ''


#funcs
def setSrcPath(file):
    global path_file1
    global path_file2
    global path_file3
    fname = QFileDialog.getOpenFileName()
    ruta = fname[0]
    nombre_archivo = os.path.basename(ruta)
    if fname[0][-5:] == ".xlsx" or  fname[0][-4:] == ".xls":
        if file == 1:
            window.txtFile1.setText(nombre_archivo)
            path_file1 = fname[0]
        elif file == 2:
            window.txtFile2.setText(nombre_archivo)
            path_file2 = fname[0]
        else:
            window.txtFile3.setText(nombre_archivo)
            path_file3 = fname[0]
    else:
        alerta = QMessageBox()
        alerta.setWindowTitle('Alerta')
        alerta.setText('Seleccione un Archivo que sea excel')
        alerta.setIcon(QMessageBox.Warning)
        alerta.setStandardButtons(QMessageBox.Ok)
        alerta.setDefaultButton(QMessageBox.Ok)
        alerta.exec()
    print(path_file1)

def processExcel():
    global path_file1
    global path_file2
    global path_file3
    if path_file1[-4:] == ".xls" or path_file1[-5:] == ".xlsx":
        libro_trabajo = xlrd.open_workbook(path_file1)
        # Especifica el índice o el nombre de la hoja que deseas utilizar
        nombre_hoja = "Calidad de Servicio Técnico"
        sheet = libro_trabajo.sheet_by_name(nombre_hoja)
        print(sheet.cell_value(13, 38))
        #print('celda 13 y 40',sheet.cell_value(13, 40))
        cal_60 = []
        for i in range(13,500):
            try:
                aux_cell = sheet.cell_value(i, 20)
                fmi = sheet.cell_value(i, 39)
                tki = sheet.cell_value(i, 40)
                empty_dict_cal60 = {}
                if aux_cell != '':
                    name_descompuesto  = re.search(r'\((.*?)\)', aux_cell).group(1)
                    empty_dict_cal60['name'] = name_descompuesto
                    empty_dict_cal60['fmik'] = fmi
                    empty_dict_cal60['ttik'] = tki
                    print(empty_dict_cal60)
                    cal_60.append(aux_cell)
            except IndexError:
                break
        print(len(cal_60))
    else:
        pass

    if path_file2[-4:] == ".xls" or path_file2[-5:] == ".xlsx":
    
        workbook = load_workbook(path_file2,data_only=True)
        nombre_hoja = "Calidad de Servicio Técnico"
        sheet_target = workbook.active
        for i in range(17,330):
            aux_cell = sheet_target[f'C{i}'].value
            aux_fmi = sheet_target[f'I{i}'].value
            aux_ttk = sheet_target[f'J{i}'].value
            empty_dict = {}
            if aux_cell != None:
                name_descompuesto  = re.search(r'\((.*?)\)', aux_cell).group(1)
                empty_dict['name'] = name_descompuesto
                empty_dict['fmik'] = aux_fmi
                empty_dict['ttik'] = aux_ttk
                print(empty_dict)
            else:
                break

    else:
        pass
    
    if path_file3[-4:] == ".xls" or path_file3[-5:] == ".xlsx":
        workbook = xlrd.open_workbook(path_file3)
        # Especifica el índice o el nombre de la hoja que deseas utilizar
       

        # Accede a la hoja deseada
        sheet = workbook.sheet_by_index(0)

        # Ahora puedes trabajar con la hoja normalmente
        # Por ejemplo, imprimir el valor de la celda en la fila 0, columna 0
        # print('8,4',sheet.cell_value(8, 4))
        # print('8,5',sheet.cell_value(8, 5))
        # print('8,6',sheet.cell_value(8, 6))
        # print('8,7',sheet.cell_value(8, 7))
        # print('8,8',sheet.cell_value(8, 8))
        # print('8,9',sheet.cell_value(8, 9))
        # print('8,10',sheet.cell_value(8, 10))
        # print('8,11',sheet.cell_value(8, 11))
        # print('8,12',sheet.cell_value(8, 12))
        # print('8,13',sheet.cell_value(8, 13))
        # print('8,14',sheet.cell_value(8, 14))
        # print('8,15',sheet.cell_value(8, 15))

        fmik_data = []
        for i in range(9,750):
            try:
                aux_cell = sheet.cell_value(i, 7)
                if aux_cell != '':
                    name_descompuesto = aux_cell.split("_")
                    aux_dic = {}
                    if len(name_descompuesto) >=2:
                        aux_dic['name'] = name_descompuesto[2].upper()
                        aux_dic['fmik'] = sheet.cell_value(9, 22)
                        aux_dic['ttik'] = sheet.cell_value(9, 23)
                        print(aux_dic)


            except IndexError:
                break

    else:
        pass
    # libro_trabajo = openpyxl.load_workbook(path_file1,data_only=True)
    # hoja = libro_trabajo['Calidad de Servicio Técnico']
    # print(hoja['C17'].value)
    # print(hoja['AG17'].value)
    # print(hoja['AH17'].value)
    # libro_trabajo.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    ui_file_name = "mainwindow.ui"
    ui_file = QFile(ui_file_name)
    if not ui_file.open(QIODevice.ReadOnly):
        print(f"Cannot open {ui_file_name}: {ui_file.errorString()}")
        sys.exit(-1)
    loader = QUiLoader()
    window = loader.load(ui_file)
    window.btnFile1.clicked.connect(lambda: setSrcPath(1))
    window.btnFile2.clicked.connect(lambda: setSrcPath(2))
    window.btnFile3.clicked.connect(lambda: setSrcPath(3))
    window.btnProcess.clicked.connect(lambda: processExcel())
    ui_file.close()
    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())


    #listado package