import os
import json
import shutil
import module

inputPath = os.getcwd() + "\input"
outputPath = os.getcwd() + "\output"
backupPath = os.getcwd()  + "\\backup"
fileNames = os.listdir(inputPath)
dictionary = json.load(open("dictionary.json", encoding='utf-8'))

qtdArquivos = fileNames.__len__()
print(str(qtdArquivos)+" ARQUIVOS ENCONTRADAS")
count = 0

for file in fileNames:
    file = file[:-5]+file[-5:].lower()
    print("ABRINDO O ARQUIVO "+file)
    shutil.copy2(inputPath+"\\"+file, outputPath+"\\classificado_"+file)
    # shutil.move(inputPath+"\\"+file, backupPath+"\\"+file)
    f = module.Excel_File(outputPath+"\\classificado_"+file)
    print("HEADER: "+str(f.header))
    print("LISTA DE PLANILHAS: "+str(f.sheet_list))
    f.read_excel.save(filename=outputPath+"\\classificado_"+file)
    count+=1
    qtdArquivos-=1
    print("PLANILHA "+str(count)+" SALVA")
    print("RESTAM "+str(qtdArquivos)+" ARQUIVOS NA FILA")