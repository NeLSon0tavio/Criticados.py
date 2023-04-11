import datetime
import email
import glob
import imaplib
import os
import shutil
import stat
import zipfile
from email.header import decode_header

import openpyxl
import pandas as pd
from dateutil.utils import today

resultadoC6 = 0
resultadopan = 0
resultadoSAFRA = 0
resultadoMERCANTIL = 0
resultadoDAYCOVAL = 0
resultadoFACTA = 0
resultado = 0
diahoje = ""
dias = 2
mes = ""
if datetime.date.today().day < 10:
    diahoje += "0"
diahoje += str(datetime.date.today().day)
if datetime.date.today().month < 10:
    mes += "0"
mes += str(datetime.date.today().month)
saida = os.getcwd() + "/Subir"
entrada = os.getcwd()
months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
Promotora = []
FindId = []
CcNumeroContrato = []
CcCliCpf = []
CcCliNome = []
CcNomeDigitadorBanco = []
CcDataStatus = []
CcStatusBanco = []


def main(exclude, baixar, salvar,dia):
    dias = dia
    #MERCANTIL(baixar)
    bmg()
    itau()
    C6(baixar)
    Facta(baixar)
    DAYCOVAL(baixar)
    SAFRA(baixar)
    pan(baixar)
    if salvar:
        data = {'Promotora': Promotora,
                'FinId': FindId,
                'CcNumeroContrato': CcNumeroContrato,
                'CcCliCpf': CcCliCpf,
                'CcCliNome': CcCliNome,
                'CcNomeDigitadorBanco': CcNomeDigitadorBanco,
                'CcDataStatus': CcDataStatus,
                'CcStatusBanco': CcStatusBanco,
                'CcDataCadastro': CcDataStatus}
        df = pd.DataFrame(data, columns=['Promotora', 'FinId', 'CcNumeroContrato',
                                         'CcCliCpf', 'CcCliNome', 'CcNomeDigitadorBanco', 'CcDataStatus',
                                         'CcStatusBanco', 'CcDataCadastro'])

        try:
            df.to_csv(saida + f'/Subir  {diahoje}{mes}.csv', index=False, sep=";", encoding="latin1")
        except OSError:
            os.mkdir(os.path.join("Subir"))
            df.to_csv(saida + f'/Subir  {diahoje}{mes}.csv', index=False, sep=";", encoding="latin1")
    if exclude:
        excluirPastas()
        
def excluirPastas():
    try:
        shutil.rmtree(os.path.join(entrada, "PAN"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "SAFRA"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "Facta"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "MERCANTIL"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "DAYCOVAL"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "C6"))
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "bmg"))
        os.mkdir("bmg")
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "itau"))
        os.mkdir("itau")
    except:
        pass
    try:
        shutil.rmtree(os.path.join(entrada, "PAN"))
    except:
        pass

def obtain_header(msg):
    dateEmail = msg["Date"]
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    return subject, dateEmail


def C6(ifbaixar):
    if ifbaixar:
        status = baixar("C6", "CRITICADOS C6")
    else:
        status = ""
    if status == "":
        caminho = entrada + '/C6'
        files_braceiro = glob.glob(caminho + "/*.csv")
        months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
        for filename in files_braceiro:
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='utf8', on_bad_lines='skip', skiprows=0, sep=';', keep_default_na=False)
            status = ""
            if str(file2.keys()[7]) != "Promotora":
                status = f"Fora do layout: Promotora -> {file2.keys()[7]}"
            if str(file2.keys()[4]) != "User Digitacao Banco":
                status += f"Fora do layout: User Digitacao Banco -> {file2.keys()[4]}"
            if str(file2.keys()[0]) != "Contrato":
                status += f"Fora do layout: Contrato -> {file2.keys()[0]}"
            if str(file2.keys()[1]) != "CPF":
                status += f"Fora do layout: CPF -> {file2.keys()[1]}"
            if str(file2.keys()[2]) != "nome":
                status += f"Fora do layout: nome -> {file2.keys()[2]}"
            if str(file2.keys()[5]) != "data status":
                status += f"Fora do layout: data status -> {file2.keys()[5]}"
            if str(file2.keys()[6]) != "Status":
                status += f"Fora do layout: Status -> {file2.keys()[6]}"
            if status == "":
                for linha in range(1, file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "" or str(file2[str(file2.keys()[1])][linha]) == "null":
                        break
                    gNumber = file2[str(file2.keys()[6])][linha]
                    if gNumber != "None" and str(gNumber) != "":
                        if "cancelada" not in str(gNumber).lower() and "reprovada" not in str(gNumber).lower() and \
                                "andamento - cadastro de proposta" not in str(gNumber).lower() and \
                                "número da ade data da digitação da ade" not in gNumber:
                            add = True
                            for i in range(len(CcNumeroContrato)):
                                if str(file2[str(file2.keys()[0])][linha]) == str(CcNumeroContrato[i]):
                                    add = False
                            if add:
                                try:
                                    year = ""
                                    date = str(file2[str(file2.keys()[5])][linha])
                                    if not date.split("/")[1].isdigit():
                                        mes = ""
                                        if months.index(date.split("/")[1]) + 1 < 10:
                                            mes += "0"
                                        mes += str(months.index(date.split("/")[1]) + 1)
                                        year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                    else:
                                        year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                            date.split("/")[0])
                                    year.replace("/", "-")
                                    a = ""
                                    if str(file2[str(file2.keys()[7])][linha]) == "" or str(
                                            file2[str(file2.keys()[7])][linha]) == '-':
                                        a = str(file2[str(file2.keys()[4])][linha])
                                    else:
                                        a = str(file2[str(file2.keys()[7])][linha])
                                    Promotora.append(a.strip())
                                    FindId.append("29")
                                    CcNumeroContrato.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                    CcCliCpf.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                    CcCliNome.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                    CcNomeDigitadorBanco.append(str(file2[str(file2.keys()[4])][linha]).strip())
                                    CcDataStatus.append(year.strip())
                                    CcStatusBanco.append(str(file2[str(file2.keys()[6])][linha]).strip())
                                except:
                                    print(f"erro no contrato {str(file2[str(file2.keys()[0])][linha]).strip()} que esta no arquivo {filename}")
            else:
                print("C6", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório
    else:
        print(status)  # fazer o relatório C6


# def MERCANTIL(ifbaixar):
#     if ifbaixar:
#         status = baixar("MERCANTIL", "CRITICADOS MERCANTIL")
#     else:
#         status = ""
#     if status == "":
#         caminho = entrada + '/MERCANTIL'
#         file = glob.glob(caminho + "/*.csv")
#         print(file)
#         for filename in file:
#             caminho2 = filename.replace("\\", "/")
#             print(filename)
#             file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';', keep_default_na=False)
#             for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
#                 if str(file2[str(file2.keys()[1])][linha]) == "nan":
#                     break
#                 gNumber = file2[str(file2.keys()[6])][linha]
#                 # if gNumber != "None" and str(gNumber) != "":
#                     # if "can" not in str(gNumber).lower() and "rep" not in str(gNumber).lower():
#                 # add = True
#                 # for i in range(len(selected)):
#                 #     if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
#                 #         add = False
#                 # if add:
#                 row_selected = [str(file2[str(file2.keys()[0])][linha]),
#                                 str(file2[str(file2.keys()[1])][linha]),
#                                 str(file2[str(file2.keys()[2])][linha]),
#                                 str(file2[str(file2.keys()[3])][linha]),
#                                 str(file2[str(file2.keys()[4])][linha]),
#                                 str(file2[str(file2.keys()[5])][linha]),
#                                 str(file2[str(file2.keys()[6])][linha]),
#                                 str(file2[str(file2.keys()[7])][linha]), "MERCANTIL",
#                                 " ",
#                                 str(file2[str(file2.keys()[8])][linha])]
#
#                 # row1.append(str(file2[str(file2.keys()[8])][linha]))
#                 print(str(file2[str(file2.keys()[0])][linha]))
#                 row2.append("16")
#                 # row3.append(str(file2[str(file2.keys()[0])][linha]).strip())
#                 # row4.append(str(file2[str(file2.keys()[1])][linha]).strip())
#                 # row5.append(str(file2[str(file2.keys()[2])][linha]).strip())
#                 # row6.append(str(file2[str(file2.keys()[3])][linha]).strip())
#                 # row7.append(str(file2[str(file2.keys()[4])][linha]).strip())
#                 # row8.append(str(file2[str(file2.keys()[5])][linha]).strip())


def Facta(ifbaixar):
    if ifbaixar:
        status = baixar("Facta", "Criticados - Robo FACTA Codigo de Promotora")
    else:
        status = ""
    if status == "":
        caminho = entrada + r"/Facta"
        files_6000041 = glob.glob(caminho + "/*.csv")
        for filename in files_6000041:
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';', keep_default_na=False)
            status = ""
            if str(file2.keys()[6]) != "Promotora":
                status = f"Fora do layout: Promotora -> {file2.keys()[6]}"
            if str(file2.keys()[0]) != "Numero Contrato":
                status += f"Fora do layout: Numero Contrato -> {file2.keys()[0]}"
            if str(file2.keys()[1]) != "CPF":
                status += f"Fora do layout: CPF -> {file2.keys()[1]}"
            if str(file2.keys()[2]) != "Nome":
                status += f"Fora do layout: Nome -> {file2.keys()[2]}"
            if str(file2.keys()[3]) != "Usuario Digitador":
                status += f"Fora do layout: Usuario Digitador -> {file2.keys()[3]}"
            if str(file2.keys()[4]) != "Data Status":
                status += f"Fora do layout: Data Status -> {file2.keys()[4]}"
            if str(file2.keys()[5]) != "Status":
                status += f"Fora do layout: Status -> {file2.keys()[5]}"
            if status == "":
                for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "" or str(file2[str(file2.keys()[1])][linha]) == "null":
                        break
                    gNumber = file2[str(file2.keys()[5])][linha]
                    if gNumber != "None":
                        if "rep" not in str(gNumber).lower() and \
                                "can" not in str(gNumber).lower():
                            add = True
                            for i in range(len(CcNumeroContrato)):
                                if str(file2[str(file2.keys()[0])][linha]) == str(CcNumeroContrato[i]):
                                    add = False
                            if add:
                                Promotora.append(str(file2[str(file2.keys()[6])][linha]).strip())
                                FindId.append("21")
                                CcNumeroContrato.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                CcCliCpf.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                CcCliNome.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                CcNomeDigitadorBanco.append(str(file2[str(file2.keys()[3])][linha]).strip())
                                CcDataStatus.append(str(file2[str(file2.keys()[4])][linha]).strip())
                                CcStatusBanco.append(str(file2[str(file2.keys()[5])][linha]).strip())
            else:
                print("Facta", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório
    else:
        print(status)  # fazer o relatório Facta


def DAYCOVAL(ifbaixar):
    if ifbaixar:
        status = baixar("DAYCOVAL", "CRITICADOS DAYCOVAL")
    else:
        status = ""
    if status == "":
        caminho = entrada + r"\DAYCOVAL"
        caminho = caminho.replace('\\', '/')
        files = glob.glob(caminho + "/*.csv")
        mes = ""
        for filename in files:
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='Utf-8', on_bad_lines='skip', skiprows=0, sep=';',keep_default_na=False )
            status = ""
            if str(file2.keys()[0]) != "Contrato":
                status = f"Fora do layout: Contrato -> {file2.keys()[0]}"
            if str(file2.keys()[1]) != "CPF":
                status = f"Fora do layout: CPF -> {file2.keys()[1]}"
            if str(file2.keys()[2]) != "nome":
                status += f"Fora do layout: nome -> {file2.keys()[2]}"
            if str(file2.keys()[4]) != "User Digitacao Banco":
                status += f"Fora do layout: User Digitacao Banco -> {file2.keys()[4]}"
            if str(file2.keys()[6]) != "Status":
                status += f"Fora do layout: Status -> {file2.keys()[6]}"
            if str(file2.keys()[5]) != "data status":
                status += f"Fora do layout: data status -> {file2.keys()[5]}"
            if str(file2.keys()[7]) != "Promotora":
                status += f"Fora do layout: Promotora -> {file2.keys()[5]}"
            if status == "":
                for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "" or str(file2[str(file2.keys()[1])][linha]):
                        break
                    gNumber = file2[str(file2.keys()[6])][linha]
                    if gNumber != "None":
                        if "rep" not in str(gNumber).lower() and \
                                "can" not in str(gNumber).lower() and \
                                "rec" not in str(gNumber).lower():
                            add = True
                            for i in range(len(CcNumeroContrato)):
                                if str(file2[str(file2.keys()[0])][linha]) == str(CcNumeroContrato[i]):
                                    add = False
                            if add:
                                year = ""
                                date = str(file2[str(file2.keys()[5])][linha])
                                if not date.split("/")[1].isdigit():

                                    if months.index(date.split("/")[1]) + 1 < 10:
                                        mes += "0"
                                    mes += str(months.index(date.split("/")[1]) + 1)
                                    year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                else:
                                    year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                        date.split("/")[0])
                                year.replace("/", "-")
                                Promotora.append(str(file2[str(file2.keys()[7])][linha]).strip())
                                FindId.append("8")
                                CcNumeroContrato.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                CcCliCpf.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                CcCliNome.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                CcNomeDigitadorBanco.append(str(file2[str(file2.keys()[4])][linha]).strip())
                                CcDataStatus.append(year.strip())
                                CcStatusBanco.append(str(file2[str(file2.keys()[6])][linha]).strip())
            else:
                print("Daycoval", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório
    else:
        print(status)  # fazer o relatório DAYCOVAL


def SAFRA(ifbaixar):
    if ifbaixar:
        status = baixar("SAFRA", "SAFRA] - CRITICADO")
    else:
        status = ""
    if status == "":
        files_normal = glob.glob(entrada + "/SAFRA/*.csv")
        months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
        for filename in files_normal:
            caminho2 = filename.replace("\\", "/")
            try:
                file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';', keep_default_na=False)
                for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "" or str(file2[str(file2.keys()[1])][linha]) == "null":
                        break
                    gNumber = file2[str(file2.keys()[7])][linha]
                    if str(gNumber) != "None" and gNumber != "+++" and gNumber != "0":
                        if "recusada" not in str(gNumber).lower() and "cancelado" not in str(gNumber).lower() and \
                                "expirado" not in str(gNumber).lower():
                            add = True
                            for i in range(len(CcNumeroContrato)):
                                if str(file2[str(file2.keys()[2])][linha]) == str(CcNumeroContrato[i]):
                                    add = False
                            if add:
                                year = ""
                                date = str(file2[str(file2.keys()[6])][linha])
                                if not date.split("/")[1].isdigit():
                                    mes = ""
                                    if months.index(date.split("/")[1]) + 1 < 10:
                                        mes += "0"
                                    mes += str(months.index(date.split("/")[1]) + 1)
                                    year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                else:
                                    year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                        date.split("/")[0])
                                year.replace("/", "-")
                                Promotora.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                FindId.append("2")
                                CcNumeroContrato.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                CcCliCpf.append(str(file2[str(file2.keys()[3])][linha]).strip())
                                CcCliNome.append(str(file2[str(file2.keys()[4])][linha]).strip())
                                CcNomeDigitadorBanco.append(str(file2[str(file2.keys()[5])][linha]).strip())
                                CcDataStatus.append(year.strip())
                                CcStatusBanco.append(str(file2[str(file2.keys()[7])][linha]).strip())
            except:
                pass

    else:
        print(status)  # fazer o relatório SAFRA


def pan(ifbaixar):
    if ifbaixar:
        status = baixar("PAN", "Relatorio de produ")
    else:
        status = ""
    if status == "":
        step = 0
        for arquivo in os.listdir(entrada + "\\Pan"):
            if arquivo.endswith('.zip'):
                try:
                    caminho_completo = os.path.join(entrada + "\\Pan", arquivo)
                    with zipfile.ZipFile(caminho_completo, 'r') as zip_ref:
                        zip_ref.extractall(entrada + "\\Pan")
                    nome_arquivo_extraido = "RelatorioProducaoCorrespondentesCompletoDIRECTX.csv"
                    novo_nome_arquivo_extraido = 'RelatorioProducaoCorrespondentesCompletoDIRECTX - ' + str(
                        step) + '.csv'
                    caminho_completo_antigo = os.path.join(entrada + "\\Pan", nome_arquivo_extraido)
                    caminho_completo_novo = os.path.join(entrada + "\\Pan", novo_nome_arquivo_extraido)
                    os.rename(caminho_completo_antigo, caminho_completo_novo)
                    os.remove(caminho_completo)

                except FileExistsError:
                    pass
            step += 1

        caminho = entrada
        files = []
        files_lewe = glob.glob(caminho + "/Pan" + "/*.csv")
        files.append(files_lewe)
        months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

        for file in files:
            for filename in file:
                caminho2 = filename.replace("\\", "/")
                file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';', keep_default_na=False)
                status = ""
                if str(file2.keys()[0]) != "ï»¿CODIGO_PROMOTORA":
                    status = f"Fora do layout: CODIGO_PROMOTORA -> {file2.keys()[0]}"
                if str(file2.keys()[20]) != "CPF_USUARIO":
                    status = f"Fora do layout: CPF_USUARIO -> {file2.keys()[20]}"
                if str(file2.keys()[1]) != "PROPOSTA":
                    status += f"Fora do layout: PROPOSTA -> {file2.keys()[1]}"
                if str(file2.keys()[21]) != "CPF_CLIENTE":
                    status += f"Fora do layout: CPF_CLIENTE -> {file2.keys()[21]}"
                if str(file2.keys()[22]) != "NOMECLI":
                    status += f"Fora do layout: NOMECLI -> {file2.keys()[22]}"
                if str(file2.keys()[10]) != "DATA_LANCAMENTO":
                    status += f"Fora do layout: DATA_LANCAMENTO -> {file2.keys()[10]}"
                if str(file2.keys()[7]) != "SITUACAO_PROPOSTA":
                    status += f"Fora do layout: SITUACAO_PROPOSTA -> {file2.keys()[7]}"
                if str(file2.keys()[9]) != "NOME_ATIVIDADE":
                    status += f"Fora do layout: NOME_ATIVIDADE -> {file2.keys()[9]}"
                if status == "":
                    for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                        if str(file2[str(file2.keys()[1])][linha]) == "" or str(file2[str(file2.keys()[1])][linha]) == "null":
                            break
                        gNumber = file2[str(file2.keys()[7])][linha] + " - " + file2[str(file2.keys()[9])][
                            linha]  # pegar somente o primeiro campo e fazer o filtro com
                        if str(gNumber) != "None" and gNumber != " " and gNumber != "0":
                            if "rep" not in str(gNumber).lower() and "can" not in str(gNumber).lower():
                                add = True
                                for x in CcNumeroContrato:
                                    if str(file2[str(file2.keys()[1])][linha]) == str(x):
                                        add = False
                                if add:
                                    year = ""
                                    date = str(file2[str(file2.keys()[6])][linha])
                                    if not date.split("/")[1].isdigit():
                                        mes = ""
                                        if months.index(date.split("/")[1]) + 1 < 10:
                                            mes += "0"
                                        mes += str(months.index(date.split("/")[1]) + 1)
                                        year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                    else:
                                        year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                            date.split("/")[0])

                                    year.replace("/", "-")
                                    Promotora.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                    FindId.append("3")
                                    CcNumeroContrato.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                    CcCliCpf.append(str(file2[str(file2.keys()[21])][linha]).strip())
                                    CcCliNome.append(str(file2[str(file2.keys()[22])][linha]).strip())
                                    CcNomeDigitadorBanco.append(
                                        str(str(file2[str(file2.keys()[20])][linha])).strip().replace(".", "").replace(
                                            "-",
                                            "") + "_00" + str(file2[str(file2.keys()[16])][linha]).strip())
                                    CcDataStatus.append(year.strip())
                                    CcStatusBanco.append(gNumber.strip())
                else:
                    print("Pan", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório
    else:
        print(status)  # fazer o relatório Pan


def baixar(finNome, filter):
    imap = imaplib.IMAP4_SSL("imap.houseti.com.br")  # establish connection

    imap.login("nelson-santos@grupoamp.com.br", "!@#AMenorParcela2020")  # login
    imap.select("inbox")
    status, response = imap.search(None, f'SUBJECT "{str(filter)}"')
    numOfMessages = response[0].split()
    date_strs = []
    ferificarUltimoEmail = True
    for i in numOfMessages[::-1]:
        e_id = str(i).replace("b'", "").replace("'", "")
        _, response = imap.fetch(e_id, "(RFC822)")
        msg = email.message_from_bytes(response[0][1])
        subject, dateEmail = obtain_header(msg)
        if "CARTAO" not in subject:
            date_format = '%d %b %Y %H:%M:%S %z'
            try:
                date_object = datetime.datetime.strptime(dateEmail, date_format)
            except ValueError:
                date_format = '%a, %d %b %Y %H:%M:%S %z'
                date_object = datetime.datetime.strptime(dateEmail, date_format)
            new_date_format = '%d %b %y'
            emaildateFormated = date_object.strftime(new_date_format)
            today = datetime.date.today()
            date_email = datetime.datetime.strptime(emaildateFormated, "%d %b %y").date()
            two_days_ago = today - datetime.timedelta(days=dias)
            if date_email < two_days_ago:
                if ferificarUltimoEmail:
                    date_strs.append(dateEmail)
                break
            else:
                ferificarUltimoEmail = False
                if msg.is_multipart():
                    for part in msg.walk():
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            namefolder = finNome
                            time = str(dateEmail[:20]).strip().replace(":", "-")
                            if finNome == "PAN":
                                filepath = os.path.join(namefolder,
                                                        part.get_filename().replace(".zip", " ") + time + " " + str(
                                                            i).replace("b'", "").replace("'", "") + ".zip")
                            else:
                                filepath = os.path.join(namefolder,
                                                        part.get_filename().replace(".csv", " ") + time + " " + str(
                                                            i).replace("b'", "").replace("'", "") + ".csv")
                            try:
                                open(filepath, "wb").write(part.get_payload(decode=True))
                            except FileNotFoundError:
                                try:
                                    os.mkdir(os.path.join(namefolder))
                                    open(filepath, "wb").write(part.get_payload(decode=True))
                                except FileNotFoundError:
                                    os.mkdir(namefolder.split("/")[0])
                                    os.mkdir(os.path.join(namefolder))
                                    open(filepath, "wb").write(part.get_payload(decode=True))
    dates = [datetime.datetime.strptime(date_str, date_format) for date_str in date_strs if
             datetime.datetime.strptime(date_str, date_format).date() < (datetime.date.today() - datetime.timedelta(
                 days=2))]
    status = ""
    if len(dates) > 0:
        status += f"O utlimo email que recebi da {finNome} foi: {max(dates).strftime('%d %b %y %H:%M:%S')}"
    imap.close()

    return status


def itau():
    caminho = entrada + '/itau'
    extrair(caminho)
    files_braceiro = glob.glob(caminho + "/*.txt")
    for filename in files_braceiro:
        caminho2 = filename.replace("\\", "/")
        with open(caminho2, "r", encoding="latin1") as arquivo:
            verificar = True

            texto = arquivo.readlines()
            for s in texto:
                spliter = ""
                if ";" in s:
                    spliter = ";"
                if "#" in s:
                    spliter = "#"
                if "§" in s:
                    spliter = "§"
                if spliter == "":
                    print("Separador não encontrado ITAU")
                else:
                    linha = s.split(spliter)  # gNumber ad(31) + dg(105)

                    status = ""
                    if verificar:
                        if str(linha[105]) != "Situação da Proposta":
                            status = f"Fora do layout: Situação da Proposta -> {linha[105]}"
                        if str(linha[30]) != "Número da ADE":
                            status += f"Fora do layout: Número da ADE -> {linha[31]}"
                        if str(linha[31]) != "Status da Operação":
                            status += f"Fora do layout: Status da Operação -> {linha[32]}"
                        if str(linha[104]) != "Usuário - inclusão":
                            status += f"Fora do layout: CPF_CLIENTE -> {linha[104]}"
                        if str(linha[7]) != "Nome do Cliente":
                            status += f"Fora do layout: Nome do Cliente -> {linha[7]}"
                        if str(linha[9]) != "CNPJ/CPF do Cliente":
                            status += f"Fora do layout: CNPJ/CPF do Cliente -> {linha[9]}"
                        if str(linha[153]) != "Data de Fator":
                            status += f"Fora do layout: Data da Digitação da ADE -> {linha[153]}"
                        if str(linha[0]) != "Loja":
                            status += f"Fora do layout: Loja -> {linha[0]}"
                        verificar = False
                    if status == "":
                        gNumber = str(linha[31] + " " + linha[105]).lower().strip()
                        add = True
                        for x in CcNumeroContrato:
                            if str(linha[30]).strip() == str(x).strip():
                                add = False
                        if add:
                            if "status" not in gNumber and "simulação de proposta" not in gNumber and "cancelad" not in gNumber and "rejeitad" not in gNumber and "reprovad" not in gNumber and "" not in gNumber and "null" not in gNumber:
                                date = str(linha[153])
                                year = date
                                if "/" in date:
                                    if not date.split("/")[1].isdigit():
                                        mes = ""
                                        if months.index(date.split("/")[1]) + 1 < 10:
                                            mes += "0"
                                        mes += str(months.index(date.split("/")[1]) + 1)
                                        year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                    else:
                                        year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                            date.split("/")[0])
                                    year.replace("/", "-")
                                Promotora.append(linha[0] + " - ITAU")
                                FindId.append("1")
                                CcNumeroContrato.append(str(linha[30]).strip())
                                CcCliCpf.append(
                                    str(linha[9]).replace(".", "").replace(".", "").replace(".", "").replace("-","").strip())
                                CcCliNome.append(str(linha[7]).strip())
                                CcNomeDigitadorBanco.append(str(linha[104]).strip())
                                CcDataStatus.append(year.strip())
                                CcStatusBanco.append(gNumber)

                    else:
                        print("ITAU", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório


def bmg():
    caminho = entrada + '/bmg'
    extrair(caminho)
    files_braceiro = glob.glob(caminho + "/*.txt")
    for filename in files_braceiro:
        caminho2 = filename.replace("\\", "/")
        with open(caminho2, "r", encoding="latin1") as arquivo:
            verificar = True
            texto = arquivo.readlines()
            for s in texto:
                spliter = ""
                if ";" in s:
                    spliter += ";"
                if "#" in s:
                    spliter += "#"
                if "§" in s:
                    spliter += "§"
                if spliter == "":
                    print("Separador não encontrado BMG")
                else:
                    linha = s.split(spliter)  # gNumber ad(31) + dg(105)
                    status = ""
                    if verificar:
                        if str(linha[110]) != "Situação da Proposta":
                            status = f"Fora do layout: Situação da Proposta -> {linha[110]}"
                        if str(linha[31]) != "Número da ADE":
                            status += f"Fora do layout: Número da ADE -> {linha[31]}"
                        if str(linha[32]) != "Status da Operação":
                            status += f"Fora do layout: Status da Operação -> {linha[32]}"
                        if str(linha[109]) != "Usuário - inclusão":
                            status += f"Fora do layout: CPF_CLIENTE -> {linha[109]}"
                        if str(linha[7]) != "Nome do Cliente":
                            status += f"Fora do layout: Nome do Cliente -> {linha[7]}"
                        if str(linha[9]) != "CNPJ/CPF do Cliente":
                            status += f"Fora do layout: CNPJ/CPF do Cliente -> {linha[9]}"
                        if str(linha[105]) != "Data da Digitação da ADE":
                            status += f"Fora do layout: Data da Digitação da ADE -> {linha[105]}"
                        if str(linha[0]) != "Loja":
                            status += f"Fora do layout: Loja -> {linha[0]}"
                        verificar = False
                    if status == "":
                        gNumber = str(linha[32] + " " + linha[110]).lower().strip()
                        if "Status da Operação" not in gNumber and "cancelad" not in gNumber and "rejeitad" not in gNumber and "reprovad" not in gNumber and "" not in gNumber and "null" not in gNumber:
                            date = str(linha[105][0:10])
                            year = date
                            if "/" in date:
                                if not date.split("/")[1].isdigit():
                                    mes = ""
                                    if months.index(date.split("/")[1]) + 1 < 10:
                                        mes += "0"
                                    mes += str(months.index(date.split("/")[1]) + 1)
                                    year += str(today().year) + "-" + mes + "-" + str(date.split("/")[0])
                                else:
                                    year += str(today().year) + "-" + date.split("/")[1] + "-" + str(
                                        date.split("/")[0])
                                year.replace("/", "-")
                            Promotora.append(linha[0] + " - BMG")
                            FindId.append("5")
                            CcNumeroContrato.append(str(linha[31]).strip())
                            CcCliCpf.append(
                                str(linha[9]).replace(".", "").replace(".", "").replace(".", "").replace("-","").strip())
                            CcCliNome.append(str(linha[7]).strip())
                            CcNomeDigitadorBanco.append(str(linha[109]).strip())
                            CcDataStatus.append(year.strip())
                            CcStatusBanco.append(gNumber)
                    else:
                        print("BMG", status, f"[Nome do arquivo] -> {filename}")  # fazer o relatório


### MERCANTIL(dia,mes)

def extrair(caminho):
    for arquivo in os.listdir(caminho):
        if arquivo.endswith('.zip'):
            try:
                caminho_completo = os.path.join(caminho, arquivo)
                with zipfile.ZipFile(caminho_completo, 'r') as zip_ref:
                    zip_ref.extractall(caminho)
                os.remove(caminho_completo)
            except FileExistsError:
                pass


main(exclude=True, baixar=True, salvar=True,dia = 5)
