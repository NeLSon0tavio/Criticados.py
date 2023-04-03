import datetime
import email
import glob
import imaplib
import os
import shutil
import stat
import zipfile
from email.header import decode_header

import pandas as pd
from dateutil.utils import today

resultadoC6 = 0
resultadopan = 0
resultadoSAFRA = 0
resultadoMERCANTIL = 0
resultadoDAYCOVAL = 0
resultadoFACTA = 0
resultado = 0
dia = ""
mes = ""
if datetime.date.today().day < 10:
    dia += "0"
dia += str(datetime.date.today().day)
if datetime.date.today().month < 10:
    mes += "0"
mes += str(datetime.date.today().month)
saida = os.getcwd() + "/Subir"
entrada = os.getcwd()
months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []
row8 = []
row9 = []


def main(exclude):
    C6()
    Facta()
    DAYCOVAL()
    SAFRA()
    pan()

    # MERCANTIL(dia, mes)
    data = {'Promotora': row1,
            'FinId': row2,
            'CcNumeroContrato': row3,
            'CcCliCpf': row4,
            'CcCliNome': row5,
            'CcNomeDigitadorBanco': row6,
            'CcDataStatus': row7,
            'CcStatusBanco': row8,
            'CcDataCadastro': row9}
    df = pd.DataFrame(data, columns=['Promotora', 'FinId', 'CcNumeroContrato',
                                     'CcCliCpf', 'CcCliNome', 'CcNomeDigitadorBanco', 'CcDataStatus',
                                     'CcStatusBanco', 'CcDataCadastro'])

    try:
        df.to_csv(saida + f'/Subir  {dia}{mes}.csv', index=False, sep=";", encoding="latin1")
    except OSError:
        os.mkdir(os.path.join("Subir"))
        df.to_csv(saida + f'/Subir  {dia}{mes}.csv', index=False, sep=";", encoding="latin1")
    if exclude:
        print("Excluindo os arquivos")
        try:
            shutil.rmtree(os.path.join(entrada, "Pan"))
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
            shutil.rmtree(os.path.join(entrada,"PAN"))
        except:
            pass
        # os.remove(entrada+"\\Daycoval")
        # os.remove(entrada+"\\Safra")
    print("terminei tudo")


def obtain_header(msg):
    dateEmail = msg["Date"]
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    return subject, dateEmail


def C6():
    print("começando c6")
    status = baixar("C6", "CRITICADOS] C6")
    if status == "":
        caminho = entrada + '/C6'
        files = []
        files_braceiro = glob.glob(caminho + "/BRACEIRO" + "/*.csv")
        files_dsa = glob.glob(caminho + "/DSA" + "/*.csv")
        files_fhbs = glob.glob(caminho + "/FHBS" + "/*.csv")
        files_jr = glob.glob(caminho + "/JR" + "/*.csv")
        files_jss = glob.glob(caminho + "/JSS" + "/*.csv")
        files_maua_dois = glob.glob(caminho + "/MAUA DOIS" + "/*.csv")
        files_silvaseg = glob.glob(caminho + "/SILVASEG" + "/*.csv")
        files_vieira = glob.glob(caminho + "/VIEIRA" + "/*.csv")
        files.append(glob.glob(caminho + "/AGUIAR" + "/*.csv"))
        files.append(files_braceiro)
        files.append(files_dsa)
        files.append(files_fhbs)
        files.append(files_jr)
        files.append(files_jss)
        files.append(files_maua_dois)
        files.append(files_silvaseg)
        files.append(files_vieira)
        months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
        selected = []
        for file in files:
            for filename in file:
                caminho2 = filename.replace("\\", "/")
                file2 = pd.read_csv(caminho2, encoding='utf8', on_bad_lines='skip', skiprows=0, sep=';')
                inicial = 0

                for linha in range(inicial, file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "nan":
                        break
                    gNumber = file2[str(file2.keys()[6])][linha]
                    if gNumber != "None" and str(gNumber) != "":
                        if "cancelada" not in str(gNumber).lower() and "reprovada" not in str(gNumber).lower() and \
                                "andamento - cadastro de proposta" not in str(gNumber).lower() and \
                                "Status" not in gNumber:
                            add = True
                            for i in range(len(selected)):
                                if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
                                    add = False
                            if add:
                                row_selected = [str(file2[str(file2.keys()[0])][linha])]
                                selected.append(row_selected)
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
                                row1.append(a.strip())
                                row2.append("29")
                                row3.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                row4.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                row5.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                row6.append(str(file2[str(file2.keys()[7])][linha]).strip())
                                row7.append(year.strip())
                                row8.append(str(file2[str(file2.keys()[6])][linha]).strip())
                                row9.append(year.strip())
    else:
        print(status)  # fazer o relatório C6
    print("fim C6")


# def MERCANTIL(dia, mes):
#     caminho = entrada + '/MERCANTIL'
#     files = [glob.glob(caminho + "/*.csv")]
#     months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
#     selected = []
#
#     for file in files:
#         for filename in file:
#             if f"{dia} {mes}" in filename:  # 9821
#                 caminho2 = filename.replace("\\", "/")
#                 file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
#                 for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
#                     if str(file2[str(file2.keys()[1])][linha]) == "nan":
#                         break
#                     gNumber = file2[str(file2.keys()[6])][linha]
#                     if gNumber != "None" and str(gNumber) != "":
#                         if "cancelada" not in str(gNumber).lower() and "reprovada" not in str(gNumber).lower() and \
#                                 "andamento - cadastro de proposta" not in str(gNumber).lower():
#                             add = True
#                             for i in range(len(selected)):
#                                 if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
#                                     add = False
#                             if add:
#                                 row_selected = [str(file2[str(file2.keys()[0])][linha]),
#                                                 str(file2[str(file2.keys()[1])][linha]),
#                                                 str(file2[str(file2.keys()[2])][linha]),
#                                                 str(file2[str(file2.keys()[3])][linha]),
#                                                 str(file2[str(file2.keys()[4])][linha]),
#                                                 str(file2[str(file2.keys()[5])][linha]),
#                                                 str(file2[str(file2.keys()[6])][linha]),
#                                                 str(file2[str(file2.keys()[7])][linha]), "MERCANTIL",
#                                                 " ",
#                                                 str(file2[str(file2.keys()[8])][linha])]
#
#                                 selected.append(row_selected)
#     bookfinal = openpyxl.Workbook()
#     bookfinal.iso_dates = True
#     sheetfinal = bookfinal.active
#
#     for x in range(0, len(selected)):
#         for coluna in sheetfinal["A" + str(x + 1) + ":K" + str(x + 1)]:
#             if "data status" not in selected[x][5]:
#                 coluna[0].value = selected[x][0]
#                 year = ""
#                 if not selected[x][5].split("/")[1].isdigit():
#                     mes = ""
#                     if months.index(selected[x][5].split("/")[1]) + 1 < 10:
#                         mes += "0"
#                     mes += str(months.index(selected[x][5].split("/")[1]) + 1)
#                     year += str(today().year) + "-" + mes + "-" + str(selected[x][5].split("/")[0])
#                 else:
#                     year += str(today().year) + "-" + selected[x][5].split("/")[1] + "-" + str(
#                         selected[x][5].split("/")[0])
#                 year.replace("/", "-")
#                 a = ""
#                 if selected[x][7] == "" or selected[x][7] == '-':
#                     a = selected[x][4]
#                 else:
#                     a = selected[x][7]
#                 row1.append(a)
#                 row2.append("16")
#                 row3.append(str(selected[x][0]))
#                 row4.append(selected[x][1])
#                 row5.append(selected[x][2])
#                 row6.append(selected[x][3])
#                 row7.append(year)
#                 row8.append(selected[x][6])
#                 row9.append(year)
#     if len(row3) != 0:
#
#         data = {'Promotora': row1,
#                 'FinId': row2,
#                 'CcNumeroContrato': row3,
#                 'CcCliCpf': row4,
#                 'CcCliNome': row5,
#                 'CcNomeDigitadorBanco': row6,
#                 'CcDataStatus': row7,
#                 'CcStatusBanco': row8,
#                 'CcDataCadastro': row9}
#         print(row1)
#         df = pd.DataFrame(data, columns=['Promotora', 'FinId', 'CcNumeroContrato',
#                                          'CcCliCpf', 'CcCliNome', 'CcNomeDigitadorBanco', 'CcDataStatus',
#                                          'CcStatusBanco', 'CcDataCadastro'])
#
#         try:
#             df.to_csv(saida + f'/Subir MERCANTIL {dia}{mes}.csv', index=False, sep=";")
#         except OSError:
#             os.mkdir(os.path.join("Subir"))
#             df.to_csv(saida + f'/Subir MERCANTIL {dia}{mes}.csv', index=False, sep=";")
#     print(f"Terminando MERCANTIL com {len(row3)} contratos")
#     return len(row3)


def Facta():
    status = baixar("Facta", "Facta")
    if status == "":
        caminho = entrada + r"/Facta"
        files = []
        files_53599 = glob.glob(caminho + "/53599" + "/*.csv")
        files_54538 = glob.glob(caminho + "/54538" + "/*.csv")
        files_92834 = glob.glob(caminho + "/92834" + "/*.csv")
        files_600053 = glob.glob(caminho + "/600053" + "/*.csv")
        files_600059 = glob.glob(caminho + "/600059" + "/*.csv")
        files_600060 = glob.glob(caminho + "/600060" + "/*.csv")
        files_600061 = glob.glob(caminho + "/600061" + "/*.csv")
        files_600062 = glob.glob(caminho + "/600062" + "/*.csv")
        files_930010 = glob.glob(caminho + "/930010" + "/*.csv")
        files_6000041 = glob.glob(caminho + "/6000041" + "/*.csv")
        files.append(files_53599)
        files.append(files_54538)
        files.append(files_92834)
        files.append(files_600053)
        files.append(files_600059)
        files.append(files_600060)
        files.append(files_600061)
        files.append(files_600062)
        files.append(files_930010)
        files.append(files_6000041)
        selected = []

        for file in files:
            for filename in file:
                caminho2 = filename.replace("\\", "/")
                file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
                for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "nan":
                        break
                    gNumber = file2[str(file2.keys()[5])][linha]
                    if gNumber != "None":
                        if "rep" not in str(gNumber).lower() and \
                                "can" not in str(gNumber).lower():
                            add = True
                            for i in range(len(selected)):
                                if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
                                    add = False
                            if add:
                                row1.append(str(file2[str(file2.keys()[6])][linha]).strip())
                                row2.append("21")
                                row3.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                row4.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                row5.append(str(file2[str(file2.keys()[2])][linha]).strip())
                                row6.append(str(file2[str(file2.keys()[3])][linha]).strip())
                                row7.append(str(file2[str(file2.keys()[4])][linha]).strip())
                                row8.append(str(file2[str(file2.keys()[5])][linha]).strip())
                                row9.append(str(file2[str(file2.keys()[7])][linha]).strip())
    else:
        print(status)  # fazer o relatório Facta


def DAYCOVAL():
    status = baixar("DAYCOVAL", "DAYCOVAL] CRITICADOS DAYCOVAL")
    if status == "":
        caminho = entrada + r"\DAYCOVAL"
        caminho = caminho.replace('\\', '/')
        files = glob.glob(caminho + "/*.csv")
        selected = []
        mes = ""
        for filename in files:
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='Utf-8', on_bad_lines='skip', skiprows=0, sep=';')
            for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                if str(file2[str(file2.keys()[1])][linha]) == "nan":
                    break
                gNumber = file2[str(file2.keys()[6])][linha]
                if gNumber != "None":
                    if "rep" not in str(gNumber).lower() and \
                            "can" not in str(gNumber).lower() and \
                            "rec" not in str(gNumber).lower():
                        add = True
                        for i in range(len(selected)):
                            if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
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
                            row1.append(str(file2[str(file2.keys()[7])][linha]).strip())
                            row2.append("8")
                            row3.append(str(file2[str(file2.keys()[0])][linha]).strip())
                            row4.append(str(file2[str(file2.keys()[1])][linha]).strip())
                            row5.append(str(file2[str(file2.keys()[2])][linha]).strip())
                            row6.append(str(file2[str(file2.keys()[3])][linha]).strip())
                            row7.append(year.strip())
                            row8.append(str(file2[str(file2.keys()[6])][linha]).strip())
                            row9.append(year.strip())
    else:
        print(status)  # fazer o relatório DAYCOVAL


def SAFRA():
    print("Começo Safra")
    status = baixar("SAFRA", "SAFRA] - CRITICADO")
    if status == "":
        files = []
        files_normal = glob.glob(entrada + "/SAFRA/NORMAL" + "/*.csv")
        files.append(files_normal)
        files_lewe = glob.glob(entrada + "/SAFRA/LEWE" + "/*.csv")
        files.append(files_lewe)
        months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
        selected = []
        for file in files:
            for filename in file:
                caminho2 = filename.replace("\\", "/")
                try:
                    file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
                    for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                        if str(file2[str(file2.keys()[1])][linha]) == "nan":
                            break
                        gNumber = file2[str(file2.keys()[7])][linha]
                        # print(gNumber)
                        if str(gNumber) != "None" and gNumber != "+++" and gNumber != "0":
                            if "recusada" not in str(gNumber).lower() and "cancelado" not in str(gNumber).lower() and \
                                    "expirado" not in str(gNumber).lower():
                                row_selected = [str(file2[str(file2.keys()[0])][linha]),
                                                str(file2[str(file2.keys()[1])][linha]),
                                                str(file2[str(file2.keys()[2])][linha]),
                                                str(file2[str(file2.keys()[3])][linha]),
                                                str(file2[str(file2.keys()[4])][linha]),
                                                str(file2[str(file2.keys()[5])][linha]),
                                                str(file2[str(file2.keys()[6])][linha]),
                                                str(file2[str(file2.keys()[7])][linha]), "SAFRA",
                                                " ",
                                                str(file2[str(file2.keys()[8])][linha])]
                                selected.append(row_selected)
                except:
                    pass

        for x in range(0, len(selected)):
            year = ""
            if not selected[x][6].split("/")[1].isdigit():
                mes = ""
                if months.index(selected[x][6].split("/")[1]) + 1 < 10:
                    mes += "0"
                mes += str(months.index(selected[x][6].split("/")[1]) + 1)
                year += str(today().year) + "-" + mes + "-" + str(selected[x][6].split("/")[0])
            else:
                year += str(today().year) + "-" + selected[x][6].split("/")[1] + "-" + str(
                    selected[x][6].split("/")[0])
            year.replace("/", "-")
            row1.append(str(selected[x][0]).strip())
            row2.append("2")
            row3.append(str(selected[x][2]).strip())
            row4.append(selected[x][3].strip())
            row5.append(selected[x][4].strip())
            row6.append(selected[x][5].strip())
            row7.append(year.strip())
            row8.append(selected[x][7].strip())
            row9.append(year.strip())
    else:
        print(status)  # fazer o relatório SAFRA
    print("fim Safra")


def pan():
    status = baixar("PAN", "Relatorio de produ")
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
                file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
                for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                    if str(file2[str(file2.keys()[1])][linha]) == "nan":
                        break

                    gNumber = file2[str(file2.keys()[7])][linha] + " - " + file2[str(file2.keys()[9])][
                        linha]  # pegar somente o primeiro campo e fazer o filtro com
                    if str(gNumber) != "None" and gNumber != " " and gNumber != "0":
                        if "rep" not in str(gNumber).lower() and "can" not in str(gNumber).lower():
                            add = True
                            for x in row3:
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
                                row1.append(str(file2[str(file2.keys()[0])][linha]).strip())
                                row2.append("3")
                                row3.append(str(file2[str(file2.keys()[1])][linha]).strip())
                                row4.append(str(file2[str(file2.keys()[20])][linha]).strip())
                                row5.append(str(file2[str(file2.keys()[22])][linha]).strip())
                                row6.append(
                                    str(str(file2[str(file2.keys()[20])][linha])).strip().replace(".", "").replace("-",
                                                                                                                   "")
                                    + "_00" + str(file2[str(file2.keys()[16])][linha]).strip())
                                row7.append(year.strip())
                                row8.append(gNumber.strip())
                                row9.append(year.strip())
    else:
        print(status)  # fazer o relatório Pan


def baixar(finNome, filter):
    status, response = imap.search(None, f'SUBJECT "{filter}"')
    numOfMessages = response[0].split()
    date_strs = []
    ferificarUltimoEmail = True
    for i in numOfMessages[::-1]:
        e_id = str(i).replace("b'", "").replace("'", "")
        _, response = imap.fetch(e_id, "(RFC822)")
        msg = email.message_from_bytes(response[0][1])
        subject, dateEmail = obtain_header(msg)
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
        two_days_ago = today - datetime.timedelta(days=2)
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
    return status


### MERCANTIL(dia,mes)

imap = imaplib.IMAP4_SSL("imap.houseti.com.br")  # establish connection

imap.login("nelson-santos@grupoamp.com.br", "!@#AMenorParcela2020")  # login

imap.select("inbox")
# C6(dia, mes) #pegar filtro
# Facta(dia, mes) #pegar filtro
# SAFRA() #pegar filtro
# DAYCOVAL()  # pronto
# pan()  #pronto
main(exclude=True)

imap.close()