#! py
######################################
#Copyright of Daniel Alejandro Rodriguez Solis, 2022#
######################################
import getopt
import os
from datetime import datetime, time
from docxtpl import DocxTemplate
from sys import argv


def getDate(dateNow):
    months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    day = dateNow.day
    month = months[dateNow.month-1]
    year = dateNow.year
    dateText = 'Tijuana B.C a {0} de {1} del {2}'.format(day, month, year)
    return dateText


def getArg():

    arguments = []
    argm = argv[1:]
    optionShort = "n:m:u:t:s:g:h"
    optionLarge = ['name=', 'matter=', 'unit=',
                   'teacher=', 'student=', 'group=', 'help=']
    try:
        opts, args = getopt.getopt(argm, optionShort, optionLarge)
    except getopt.GetoptError as error:
        print('Se ha generado un error\n INFORMACION DEL ERROR: {}'.format(error))
        opts = []

    for opt, arg in opts:
        if opt == '-n':
            arguments.append(arg)
        if opt == '-m':
            arguments.append(arg)
        if opt == '-u':
            arguments.append(arg)
        if opt == '-t':
            arguments.append(arg)
        if opt == '-s':
            arguments.append(arg)
        if opt == '-g':
            arguments.append(arg)
        if opt == '-h':
            greeting()

    return arguments


def readData(data):

    if len(data) < 5:
      print('La cantidad de parametros enviados no cumplen con lo minimo')
    else:
        template = DocxTemplate(
            r'D:\dany1\Documents\octavoSemestre\PortadaTecnologico.docx')
        now = datetime.now()
        logTime = getDate(now)
        text = {
            "NombreTarea": data[0],
            'NombreMateria': data[1],
            'Unidad': data[2],
            'Maestro': data[3],
            'Alumno': data[4],
            'Grupo': data[5],
            'Fecha': logTime
        }
        createFolder(data[1])
        newFile = r'D:\dany1\Documents\octavoSemestre\{0}\{1}-{2}.docx'.format(
            data[1], data[4], data[0])

        template.render(text)
        if os.path.exists(newFile):
            print('Esta portada ya existe\n Desea guardar de todas formas?\n1.Si\n2.No')
            res = input()
            if int(res) == 1:
                daate = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
                newFileDoc = r'D:\dany1\Documents\octavoSemestre\{0}\{1}-{2}-{3}.docx'.format(
                    data[1], data[4], data[0], daate)
                template.save(newFileDoc)
            else:
                exit()
        else:
            template.save(newFile)


def createFolder(folderName):
    if os.path.exists(folderName) == True:
        pass
    else:
        os.mkdir(folderName)


def greeting():
    print()
    print(" |           |    _____         __        _____            ___    _            ______       ")
    print(" | |       | |   |     \       /  \      |     \    |  /    |    | \     |    /            ")
    print("  \ \ |_| / /    |      |     /    \     |     |    | /     |    |  \    |   |            ")
    print("    \ | | /      |      |    /      \    |_____|    |/      |    |   \   |   |  ______     ")
    print("     /|_|\       |      |   /________\   |     \    |\      |    |    \  |   |        |     ")
    print("   / /   \ \     |      |  /          \  |      \   | \     |    |     \ |   |        |  ")
    print("   | |    | |    |_____/  /            \ |       \  |  \   _|_   |      \|   |________|            ")
    print('____________________________________________________________________________________________________')
    print('____________________________________________________________________________________________________')
    print('[-] Bienvenido este es un programa cuyo proposito es:')
    print('[-] Crear portadas para tus materias de una manera mucho mas facil y rapida, solo se necesita pasar los parametros necesarios')
    print('[-] INFORMACION PARAMETROS:')
    print(
        '[+] -h --> Invoca la bienvenida junto con la informacion de los parametros ')
    print('[+] -n --> Es el nombre del trabajo/proyecto/tarea que ira en la portada')
    print('[+] -m --> Es el nombre de la materia ')
    print('[+] -u --> Es el nombre de unidad')
    print('[+] -t --> Es el nombre del maestro')
    print('[+] -s --> Es el nombre del alumno')
    print('[+] -g --> Es el nombre del grupo ')


def main():
    insertArgument = getArg()
    if len(insertArgument) == 0:
        pass
    else:
        readData(insertArgument)


if __name__ == "__main__":
    main()
