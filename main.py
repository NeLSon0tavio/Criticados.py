import glob
import threading
import datetime
import imaplib
import email

import openpyxl
import pandas as pd
from email.header import decode_header
import os

from dateutil.utils import today

import CriticadosC6, CriticadosFacta, CriticadosDaycoval

imap = imaplib.IMAP4_SSL("imap.houseti.com.br")  # establish connection

imap.login("nelson-santos@grupoamp.com.br", "!@#AMenorParcela2020")  # login

# print(imap.list())  # print various inboxes
status, messages = imap.select("INBOX")  # select inbox

numOfMessages = int(messages[0])  # get number of messages
row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []
row8 = []
row9 = []
saida = os.getcwd() + "/Subir"

entrada = os.getcwd()
dia = ""
mes = ""
if datetime.date.today().day < 10:
    dia += "0"
dia += str(datetime.date.today().day)
if datetime.date.today().month < 10:
    mes += "0"
mes += str(datetime.date.today().month)


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def obtain_header(msg):
    dateEmail = msg["Date"]
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    return subject, dateEmail


def download_attachment(subject, part, dateEmail):
    filename = part.get_filename()

    if filename:
        namefolder = ""
        if "MERCANTIL] CRITICADOS" in subject:
            namefolder = "Mercantil"
        if "FACTA" in filename:
            namefolder = "FACTA"
        if "criticados cartao" in filename:
            namefolder = "DAYCOVAL"
        if "criticados.csv" == filename and "C6" in subject:
            namefolder = "C6"
            if "JSS" in subject:
                namefolder += "/JSS"
            if "VIEIRA" in subject:
                namefolder += "/VIEIRA"
            if "SILVASEG" in subject:
                namefolder += "/SILVASEG"
            if "BRACEIRO" in subject:
                namefolder += "/BRACEIRO"
            if "MAUA DOIS" in subject:
                namefolder += "/MAUA DOIS"
            if "JR" in subject:
                namefolder += "/JR"
            if "DSA" in subject:
                namefolder += "/DSA"
            if "AGUIAR" in subject:
                namefolder += "/AGUIAR"
            if "FHBS" in subject:
                namefolder += "/FHBS"
        if "FACTA" in filename:
            namefolder = "FACTA"
            if "6000041" in subject:
                namefolder += "/6000041"
            if "600061" in subject:
                namefolder += "/600061"
            if "930010" in subject:
                namefolder += "/930010"
            if "600060" in subject:
                namefolder += "/600060"
            if "600062" in subject:
                namefolder += "/600062"
            if "600059" in subject:
                namefolder += "/600059"
            if "92834" in subject:
                namefolder += "/92834"
            if "53599" in subject:
                namefolder += "/53599"
            if "54538" in subject:
                namefolder += "/54538"
            if "600053" in subject:
                namefolder += "/600053"
        folder_name = clean(subject)
        if namefolder != "":
            time = str(dateEmail[4:22]).strip().replace(":", "-")
            filename = filename.replace(".csv", " ") + time + ".csv"
            if not os.path.isdir(folder_name):
                filepath = os.path.join(namefolder, filename)
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


for i in range(numOfMessages, numOfMessages - 1000, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            try:
                subject, dateEmail = obtain_header(msg)
            except:
                continue
            if msg.is_multipart():
                for part in msg.walk():
                    body = ""
                    charset = part.get_content_charset()
                    if part.get_content_type() == "text/plain":
                        partStr = part.get_payload(decode=True)
                        body += partStr.decode(charset)
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    #                    if "SAFRA" in subject and "CRITICADO" in subject:
                    #                        body = part.get_payload(decode=True)
                    #                        body = str(body).replace("b'Usuario: NORMAL!'", "").replace("b'Usuario: LEWE!'", "").replace(
                    #                            "b'<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">", "")
                    #                        text = str(body)
                    #                        linhas = text.split("<br>")
                    #
                    #                        if text != "None" and linhas[0] != "":
                    #                            for linha in linhas:
                    #                                # gNumber = linha[6]
                    #                                # if gNumber != "None":
                    #                                #     result = 1
                    #                                #     if "Can" in gNumber or \
                    #                                #             "Rep" in gNumber or \
                    #                                #             "Rec" in gNumber:
                    #                                #         result = 0
                    #                                #     if result == 1:
                    #                                # row_selected = [str(file2[str(file2.keys()[0])]),
                    #                                #                 str(file2[str(file2.keys()[1])]),
                    #                                #                 str(file2[str(file2.keys()[2])]),
                    #                                #                 str(file2[str(file2.keys()[3])]).replace("nan", ""),
                    #                                #                 str(file2[str(file2.keys()[4])]).replace("nan", ""),
                    #                                #                 str(file2[str(file2.keys()[5])]),
                    #                                #                 str(file2[str(file2.keys()[6])]),
                    #                                #                 str(file2[str(file2.keys()[7])]), "Daycoval",
                    #                               #                 "",
                    #                                #                 str(file2[str(file2.keys()[10])])]
                    #                                # selected.append(row_selected)
                    #                                # contrato/ CPF/ nome/ Vendedor/ usuarioDigitadorBanco/ data/ staus/ promotra/ banco/ dt cad
                    #                                linha.split(";")[2]  # contrato C
                    #                                linha.split(";")[3]  # CPF D
                    #                                linha.split(";")[4]  # Nome E
                    #                                linha.split(";")[5]  # vendedor D
                    #                                linha.split(";")[5]  # usuarioDigitadorBanco F
                    #                                linha.split(";")[6]  # data G
                    #                                linha.split(";")[7]  # status H
                    #                                linha.split(";")[0]  # promotora A
                    #                                banco = "SAFRA"  # banco
                    #                                dePara = ""
                    #                                r8 = linha.split(";")[7]  # data I
                    #                                # ss = bytes(r8, "latin-1").decode()
                    #                                # print(ss.encode("utf-8"))
                    #                                # print(r8.encode("utf-8", "replace"))
                    #
                    #                            promotora = ""
                    #                            if linha.split(";")[0] == "CAPITAL 2":
                    #                                promotora = linha.split(";")[5]
                    #                            else:
                    #                                promotora = linha.split(";")[0]
                    #                            if linha.split(";")[7] != "+++" and linha.split(";")[7] != "0" and "recusada" not in \
                    #                                    linha.split(";")[7] and "cancelado" not in linha.split(";")[
                    #                                7] and "expirado" not in linha.split(";")[7]:
                    #                                row1.append(promotora)
                    #                                row2.append("21")
                    #                                r3 = linha.split(";")[2]
                    #                                r4 = linha.split(";")[3]
                    #                                r5 = linha.split(";")[2]
                    #                                r6 = linha.split(";")[4]
                    #                                r7 = linha.split(";")[6]
                    #                                sss = "Em Digita\xc3\xa7\xc3\xa3oATIVO"
                    #                                r8 = linha.split(';')[7]
                    #                                # print(r8.encode("utf-8").decode("utf-8"))
                    #                                # for s in r8:
                    #                                #     print(s.encode("latin-1").decode("utf-8"))
                    #                                # print(r8)
                    #                                # r8 = r8.encode('latin-1').decode('utf-8')
                    #                                # print(r8.encode("ascii").decode("unicode_escape").encode(
                    #                                #     'Latin-1').decode('utf-8'))
                    #                                row3.append(r3.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                row4.append(r4.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                row5.append(r5.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                row6.append(r6.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                row7.append(r8.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                row8.append(r8.encode("ascii").decode("unicode_escape").encode(
                    #                                    'Latin-1').decode('utf-8'))
                    #                                year = linha.split(";")[8].split("/")[2] + "-" + linha.split(";")[8].split("/")[
                    #                                    1] + "-" + linha.split(";")[8].split("/")[0]
                    #                                row9.append(year)
                    #                    else:
                    if "attachment" in content_disposition:
                        download_attachment(subject, part, dateEmail)
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
                                 'CcCliCpf', 'CcCliNome', 'CcNomeDigitadorBanco',
                                 'CcDataStatus',
                                 'CcStatusBanco', 'CcDataCadastro'])

# try:
#     df.to_csv(saida + f'/Subir SAFRA {dia}{mes}.csv', index=False, sep=";")
# except OSError:
#     os.mkdir(os.path.join("Subir"))
#     df.to_csv(saida + f'/Subir SAFRA {dia}{mes}.csv', index=False, sep=";")

# else:
# content_type = msg.get_content_type()
# body = msg.get_payload(decode=True).decode()

imap.close()

dia = ""
mes = ""

if datetime.date.today().day < 10:
    dia += "0"
dia += str(datetime.date.today().day)
if datetime.date.today().month < 10:
    mes += "0"
mes += str(datetime.date.today().month)

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

row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []
row8 = []
row9 = []

for file in files:
    for filename in file:
        if f"{dia} {months[datetime.date.today().month - 1].title()}" in filename:  # 9821
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
            for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                if str(file2[str(file2.keys()[1])][linha]) == "nan":
                    break
                gNumber = file2[str(file2.keys()[6])][linha]
                if gNumber != "None":
                    if "CANCELADA" not in gNumber and "REPROVADA" not in gNumber and \
                            "ANDAMENTO - CADASTRO DE PROPOSTA" not in gNumber:
                        add = True
                        for i in range(len(selected)):
                            if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
                                add = False
                        if add:
                            row_selected = [str(file2[str(file2.keys()[0])][linha]),
                                            str(file2[str(file2.keys()[1])][linha]),
                                            str(file2[str(file2.keys()[2])][linha]),
                                            str(file2[str(file2.keys()[3])][linha]),
                                            str(file2[str(file2.keys()[4])][linha]),
                                            str(file2[str(file2.keys()[5])][linha]),
                                            str(file2[str(file2.keys()[6])][linha]),
                                            str(file2[str(file2.keys()[7])][linha]), "C6 consig",
                                            " ",
                                            str(file2[str(file2.keys()[8])][linha])]

                            selected.append(row_selected)
bookfinal = openpyxl.Workbook()
bookfinal.iso_dates = True
sheetfinal = bookfinal.active

for x in range(1, len(selected)):
    for coluna in sheetfinal["A" + str(x + 1) + ":K" + str(x + 1)]:
        coluna[0].value = selected[x][0]
        year = ""
        if not selected[x][5].split("/")[1].isdigit():
            mes = ""
            if months.index(selected[x][5].split("/")[1]) + 1 < 10:
                mes += "0"
            mes += str(months.index(selected[x][5].split("/")[1]) + 1)
            year += str(today().year) + "-" + mes + "-" + str(selected[x][5].split("/")[0])
        else:
            year += str(today().year) + "-" + selected[x][5].split("/")[1] + "-" + str(
                selected[x][5].split("/")[0])
        year.replace("/", "-")
        a = ""
        if selected[x][7] == "" or selected[x][7] == '-':
            a = selected[x][4]
        else:
            a = selected[x][7]
        coluna[0].value = a
        coluna[1].value = "29"
        coluna[2].value = selected[x][0]
        coluna[3].value = selected[x][1]
        coluna[4].value = selected[x][2]
        coluna[5].value = selected[x][3]
        coluna[6].value = year
        coluna[7].value = selected[x][6]
        coluna[8].value = year
        row1.append(a)
        row2.append("29")
        row3.append(str(selected[x][0]))
        row4.append(selected[x][1])
        row5.append(selected[x][2])
        row6.append(selected[x][3])
        row7.append(year)
        row8.append(selected[x][6])
        row9.append(year)
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
    df.to_csv(saida + f'/Subir C6 {dia}{mes}.csv', index=False, sep=";")
except OSError:
    os.mkdir(os.path.join("Subir"))
    df.to_csv(saida + f'/Subir C6 {dia}{mes}.csv', index=False, sep=";")
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
months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
selected = []

row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []
row8 = []
row9 = []

for file in files:
    for filename in file:
        if f"{dia} {months[datetime.date.today().month - 1].title()}" in filename:  # 9821
            caminho2 = filename.replace("\\", "/")
            file2 = pd.read_csv(caminho2, encoding='latin-1', on_bad_lines='skip', skiprows=0, sep=';')
            for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
                if str(file2[str(file2.keys()[1])][linha]) == "nan":
                    break
                gNumber = file2[str(file2.keys()[5])][linha]
                if gNumber != "None":
                    if not gNumber.find("CANCELADO") and \
                            not gNumber.find("RROVADO") and \
                            not gNumber.find("Rec"):
                        add = True
                        for i in range(len(selected)):
                            if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
                                add = False
                        if add:
                            row_selected = [str(file2[str(file2.keys()[0])][linha]),
                                            str(file2[str(file2.keys()[1])][linha]),
                                            str(file2[str(file2.keys()[2])][linha]),
                                            str(file2[str(file2.keys()[3])][linha]),
                                            str(file2[str(file2.keys()[4])][linha]),
                                            str(file2[str(file2.keys()[5])][linha]),
                                            str(file2[str(file2.keys()[6])][linha]),
                                            str(file2[str(file2.keys()[7])][linha]),
                                            "Facta",
                                            " ",
                                            str(file2[str(file2.keys()[7])][linha])]
                            selected.append(row_selected)

for x in range(1, len(selected)):
    row1.append(selected[x][3])
    row2.append("29")
    row3.append(str(selected[x][0]))
    row4.append(selected[x][1])
    row5.append(selected[x][2])
    row6.append(selected[x][3])
    row7.append(selected[x][4])
    row8.append(selected[x][6])
    row9.append(selected[x][4])
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
df.to_csv(saida + f'/Subir Facta {dia}{mes}.csv', index=False, sep=";")

caminho = entrada + r"\DAYCOVAL"
caminho = caminho.replace('\\', '/')
months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
files = glob.glob(caminho + "/*.csv")

selected = []

row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []
row8 = []
row9 = []

for filename in files:
    if f"{dia} {months[datetime.date.today().month - 1].title()}" in filename:  # 9821
        caminho2 = filename.replace("\\", "/")
        file2 = pd.read_csv(caminho2, encoding='Utf-8', on_bad_lines='skip', skiprows=0, sep=';')
        for linha in range(file2[str(file2.keys()[1])].size):  # mudar para coluna
            if str(file2[str(file2.keys()[1])][linha]) == "nan":
                break
            gNumber = file2[str(file2.keys()[6])][linha]
            if gNumber != "None":
                if "Can" not in gNumber and \
                        "Rep" not in gNumber and \
                        "Rec" not in gNumber:
                    add = True
                    for i in range(len(selected)):
                        if str(file2[str(file2.keys()[0])][linha]) == str(selected[i][0]):
                            add = False
                    if add:
                        row_selected = [str(file2[str(file2.keys()[0])][linha]),
                                        str(file2[str(file2.keys()[1])][linha]),
                                        str(file2[str(file2.keys()[2])][linha]),
                                        str(file2[str(file2.keys()[3])][linha]).replace("nan", ""),
                                        str(file2[str(file2.keys()[4])][linha]).replace("nan", ""),
                                        str(file2[str(file2.keys()[5])][linha]),
                                        str(file2[str(file2.keys()[6])][linha]),
                                        str(file2[str(file2.keys()[7])][linha]), "Daycoval",
                                        "",
                                        str(file2[str(file2.keys()[10])][linha])]

                        selected.append(row_selected)

for x in range(0, len(selected)):
    year = ""
    if not selected[x][5].split("/")[1].isdigit():
        mes = ""
        if months.index(selected[x][5].split("/")[1]) + 1 < 10:
            mes += "0"
        mes += str(months.index(selected[x][5].split("/")[1]) + 1)
        year += str(today().year) + "-" + mes + "-" + str(selected[x][5].split("/")[0])
    else:
        year += str(today().year) + "-" + selected[x][5].split("/")[1] + "-" + str(
            selected[x][5].split("/")[0])
    year.replace("/", "-")
    row1.append(selected[x][7])
    row2.append("8")
    row3.append(selected[x][0])
    row4.append(selected[x][1])
    row5.append(selected[x][2])
    row6.append(selected[x][3])
    row7.append(year)
    row8.append(selected[x][6])
    row9.append(year)
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
df.to_csv(saida + f'/Subir Daycoval {dia}{mes}.csv', index=False, sep=";")
