
import os
import ssl
import sys
import smtplib
import openpyxl
import xlsxwriter

from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#-----------------------------------------------------------------------------------------------------------------------
# Criação de arquivo Excel (relatório).

# Cria um objeto Workbook, que é o documento de Excel.
workbook = xlsxwriter.Workbook("Relatório.xlsx")

# Cria a planilha Relatório dentro do Workbook.
worksheet = workbook.add_worksheet("Relatório")

# Ajusta as larguras das colunas em uma escala de 1 para 7,4.
worksheet.set_column("C:C", 40) # +- 300
worksheet.set_column("D:D", 67) # +- 500

worksheet.set_column("E:E", 8) # +- 60
worksheet.set_column("F:F", 12) # +- 90
worksheet.set_column("G:G", 8) # +- 60
worksheet.set_column("H:H", 12) # +- 90

# Cria cabeçalho do relatório.
worksheet.write("A1", "Enviado")
worksheet.write("B1", "Código")
worksheet.write("C1", "Email")
worksheet.write("D1", "Empresa")
worksheet.write("E1", "DAS")
worksheet.write("F1", "Faturamento")
worksheet.write("G1", "DARF")
worksheet.write("H1", "Parcelamento")

# Índice ascendente associado as linhas do excel sendo lido.
indice_excel = 1

#-----------------------------------------------------------------------------------------------------------------------
# Leitura de arquivo Excel (lista de clientes).

# Acha o arquivo.
nome_arquivo_excel = None
todos_os_arquivos = os.listdir()

for arquivo in todos_os_arquivos:
    if arquivo.endswith(".xlsx"):
        nome_arquivo_excel = arquivo

# Importa o arquivo.
documento = openpyxl.load_workbook(nome_arquivo_excel)

# Carrega a primeira planilha do arquivo.
sheet = documento.worksheets[0]

# Cria listas com os valores separados.
col_codigos = [linha[0] for linha in sheet.values] # Dados da coluna 1 do excel
col_emails = [linha[1] for linha in sheet.values] # Dados da coluna 2 do excel
col_nomes = [linha[2] for linha in sheet.values] # Dados da coluna 3 do excel

#-----------------------------------------------------------------------------------------------------------------------
# TESTE: Checa se o tamanho de todas as listas é equivalente.

# print(f"{len(col_codigos)} | {len(col_emails)} | {len(col_nomes)}")

#-----------------------------------------------------------------------------------------------------------------------
# TESTE: Seleciona alguns clientes aleatórios para checar o alinhamento de tabelas.

# for i in range(5):
#    cliente_aleatorio = random.randint(0, len(col_codigos))
#    print(f"{col_codigos[cliente_aleatorio]} | {col_emails[cliente_aleatorio]} | {col_nomes[cliente_aleatorio]}")

#-----------------------------------------------------------------------------------------------------------------------
# Construção do servidor.

# Dados para o servidor.
email_origem = "email@email.com"
senha_do_app = "xxxx xxxx xxxx xxxx"

# Cria um contexto SSL seguro
contexto = ssl.create_default_context()

# Cria o servidor.
servidor = smtplib.SMTP("smtp.gmail.com", 587)
servidor.starttls(context=contexto)
servidor.login(email_origem, senha_do_app)

#-----------------------------------------------------------------------------------------------------------------------
# Pergunta qual o tipo do email.
tipo_email = input("Tipo do email. Digite 1 para 'Normal', 2 para 'Lembrete' ou 3 para 'Parcelamento': ")

# Confirma a validade do tipo do email fornecido.
try:
    if tipo_email not in ["1", "2", "3"]:
        raise ValueError()

# Fecha o programa caso o tipo fornecido esteja errado.
except ValueError:
    print("TIPO INVÁLIDO, por favor reinicie o programa.")
    sys.exit()

#-----------------------------------------------------------------------------------------------------------------------
# Checa a validade dos dados fornecidos relativos ao tipo de email desejado.

# Cria variáveis vazias para prevenir erros ao criar o email.
data_boleto_DAS = None
data_boleto_DARF = None
data_boleto_unica = None # Exclusiva para Parcelamentos.
data_para_assunto = None # Usada no "assunto" do email.

# Caso seja um email NORMAL ou LEMBRETE.
if int(tipo_email) < 3:

    # Pede as datas de vencimento.
    data_boleto_DAS = input("Data de vencimento DAS no formato DD/MM/AAAA: ")
    data_boleto_DARF = input("Data de vencimento DARF no formato DD/MM/AAAA: ")

    # Confirma a validade e formatação das datas fornecidas.
    try:
        datetime.strptime(data_boleto_DAS, "%d/%m/%Y")
        datetime.strptime(data_boleto_DARF, "%d/%m/%Y")

    # Fecha o programa caso as datas fornecidos sejam inválidas.
    except ValueError:
        print("DATA(S) INVÁLIDA(S), por favor reinicie o programa.")
        sys.exit()

# Caso seja um email de PARCELAMENTO.
else:
    
    # Pede a data de vencimento.
    data_boleto_unica = input("Data de vencimento no formato DD/MM/AAAA: ")

    # Confirma a validade e formatação das datas fornecidas.
    try:
        datetime.strptime(data_boleto_unica, "%d/%m/%Y")

    # Fecha o programa caso as datas fornecidos sejam inválidas.
    except ValueError:
        print("DATA INVÁLIDA, por favor reinicie o programa.")
        sys.exit()

#-----------------------------------------------------------------------------------------------------------------------
# Variáveis para registro de listas.

# Prepara listas com códigos de cliente que tem PDF correspondente para envio de email NORMAL ou LEMBRETE.
pastas_normal_lembrete = ["DAS", "Faturamentos", "FOLHA"]

# Listas para códigos.
lista_cod_DAS = []
lista_cod_faturamentos = []
lista_cod_DARF = []
lista_cod_parcelamentos = [] # PODE CONTER MAIS DE 1 RESULTADO POR CLIENTE!

# Listas para nome de arquivos.
lista_DAS = []
lista_faturamentos = []
lista_DARF = []
lista_parcelamentos = [] # PODE CONTER MAIS DE 1 RESULTADO POR CLIENTE!

# Se o email NÃO FOR de parcelamento!
# Procura as pastas DAS, Faturamentos e FOLHA entre todos os arquivos da pasta.
if int(tipo_email) < 3:

    for item in pastas_normal_lembrete:

        # Define qual pasta está sendo usada.
        pasta = os.listdir(item)

        # Popula listas de códigos de cliente e seus respectivos nomes de documento.
        for nome_arquivo in pasta:

            if item == "DAS":
                codigo_cliente = nome_arquivo.split("_")[1]
                lista_cod_DAS.append(int(codigo_cliente))
                lista_DAS.append(nome_arquivo)

            elif item == "Faturamentos":
                codigo_cliente = nome_arquivo.split(" - ")[1]
                lista_cod_faturamentos.append(int(codigo_cliente))
                lista_faturamentos.append(nome_arquivo)

            elif item == "FOLHA":
                posicao = nome_arquivo.find(" - ")
                codigo_cliente = nome_arquivo[:posicao]
                tipo_documento = nome_arquivo[posicao + 3:posicao + 4]
                
                if tipo_documento == "D":
                    lista_cod_DARF.append(int(codigo_cliente))
                    lista_DARF.append(nome_arquivo)

# Procura a pasta Parcelamentos.
else:

    # Popula listas de códigos de cliente e seus respectivos nomes de documento.
    for nome_arquivo in os.listdir("Parcelamentos"):

        posicao = nome_arquivo.find(" - ")
        codigo_cliente = nome_arquivo[:posicao]

        lista_cod_parcelamentos.append(int(codigo_cliente))
        lista_parcelamentos.append(nome_arquivo)

#-----------------------------------------------------------------------------------------------------------------------
# TESTE: Mostra códigos de cliente e seus arquivos associados.

# for i in range(len(lista_DAS)):
#     print(f"DAS: {lista_cod_DAS[i]} | {lista_DAS[i]}")

# for i in range(len(lista_faturamentos)):
#     print(f"FAT: {lista_cod_faturamentos[i]} | {lista_faturamentos[i]}")

# for i in range(len(lista_DARF)):
#     print(f"DARF: {lista_cod_DARF[i]} | {lista_DARF[i]}")

# for i in range(len(lista_parcelamentos)):
#     print(f"DARF: {lista_cod_parcelamentos[i]} | {lista_parcelamentos[i]}")

#-----------------------------------------------------------------------------------------------------------------------
# Cria e envia emails.

# Executa uma vez para cada código de cliente (lembrando que um código pode ter múltiplos emails).
for fileira, codigo_cliente in enumerate(col_codigos):

    # Define nome da empresa.
    nome_empresa = col_nomes[fileira]

    # Cria uma lista com todos os emails, mesmo que seja um email só.
    todos_os_emails = col_emails[fileira].split(";")

    # Cria uma mensagem para cada email na lista criada acima.
    for email_destino in todos_os_emails:

        #---------------------------------------------------------------------------------------------------------------
        # Processo de criação de emails.

        # Atualiza o indice para ir para a próxima linha do excel.
        indice_excel += 1

        # Marca no relatório a identidade do destinatário.
        worksheet.write("B" + str(indice_excel), codigo_cliente)
        worksheet.write("C" + str(indice_excel), email_destino)
        worksheet.write("D" + str(indice_excel), nome_empresa)

        #---------------------------------------------------------------------------------------------------------------
        # Define a data que aparece no título do email.

        # Se o email NÃO FOR de parcelamento!
        if int(tipo_email) < 3:
    
            if data_boleto_DAS == data_boleto_DARF:
                data_para_assunto = data_boleto_DAS

            else:
                data_para_assunto = f"{data_boleto_DAS} e {data_boleto_DARF}"
        
        # Email de parcelamento.
        else:
            data_para_assunto = data_boleto_unica

        #---------------------------------------------------------------------------------------------------------------
        # Define o texto que será enviado dependendo do tipo de email.

        # Cria a mensagem (objeto do tipo MIME), para poder aceitar HTML e anexos.
        mensagem = MIMEMultipart("alternative")
        mensagem["Subject"] = "Assunto do email"
        mensagem["From"] = email_origem
        mensagem["To"] = email_destino

        #---------------------------------------------------------------------------------------------------------------
        # Texto de suporte para quem não usa um cliente de email que lê HTML (extremamente raro).
        texto_erro = "Por gentileza, acesse sua conta por um cliente de email que leia HTML, como Gmail, Outlook, etc."

        # Texto comum de envio.
        texto_para_envio = "Esse é o texto HTML."

        #---------------------------------------------------------------------------------------------------------------
        # Cria e anexa os elementos de texto do email.
        texto_mime_erro = MIMEText(texto_erro, "plain")
        texto_mime_html = MIMEText(texto_para_envio, "html")

        mensagem.attach(texto_mime_erro)
        mensagem.attach(texto_mime_html)

        #---------------------------------------------------------------------------------------------------------------
        # Se o email NÃO FOR de parcelamento!
        # Anexa DAS se houver PDF disponível para o código do cliente.
        if (int(tipo_email) < 3) & (codigo_cliente in lista_cod_DAS):

            # Marca no relatório que houve DAS.
            worksheet.write("E" + str(indice_excel), "✔️")

            # Encontra o nome do arquivo, com o índice do código.
            indice_lista = lista_cod_DAS.index(codigo_cliente)
            nome_arquivo_pdf = "DAS/" + lista_DAS[indice_lista]

            # Importa PDF para anexação posterior.
            with open(nome_arquivo_pdf, "rb") as arquivo:
                arquivo_mime = MIMEApplication(arquivo.read())
                arquivo_mime.add_header("Content-Disposition", "attachment", filename=nome_arquivo_pdf)
                mensagem.attach(arquivo_mime)

        # Caso não tenha DAS.
        else:
            worksheet.write("E" + str(indice_excel), "❌")

        #---------------------------------------------------------------------------------------------------------------
        # Se o email NÃO FOR de parcelamento!
        # Anexa Faturamento se houver PDF disponível para o código do cliente.
        if (int(tipo_email) < 3) & (codigo_cliente in lista_cod_faturamentos):

            # Marca no relatório que houve Faturamento.
            worksheet.write("F" + str(indice_excel), "✔️")

            # Encontra o nome do arquivo, com o índice do código.
            indice_lista = lista_cod_faturamentos.index(codigo_cliente)
            nome_arquivo_pdf = "Faturamentos/" + lista_faturamentos[indice_lista]

            # Importa PDF para anexação posterior.
            with open(nome_arquivo_pdf, "rb") as arquivo:
                arquivo_mime = MIMEApplication(arquivo.read())
                arquivo_mime.add_header("Content-Disposition", "attachment", filename=nome_arquivo_pdf)
                mensagem.attach(arquivo_mime)

        # Caso não tenha Faturamento.
        else:
            worksheet.write("F" + str(indice_excel), "❌")

        #---------------------------------------------------------------------------------------------------------------
        # Se o email NÃO FOR de parcelamento!
        # Anexa DARF se houver PDF disponível para o código do cliente.
        if (int(tipo_email) < 3) & (codigo_cliente in lista_cod_DARF):

            # Marca no relatório que houve DARF.
            worksheet.write("G" + str(indice_excel), "✔️")

            # Encontra o nome do arquivo, com o índice do código.
            indice_lista = lista_cod_DARF.index(codigo_cliente)
            nome_arquivo_pdf = "FOLHA/" + lista_DARF[indice_lista]

            # Importa PDF para anexação posterior.
            with open(nome_arquivo_pdf, "rb") as arquivo:
                arquivo_mime = MIMEApplication(arquivo.read())
                arquivo_mime.add_header("Content-Disposition", "attachment", filename=nome_arquivo_pdf)
                mensagem.attach(arquivo_mime)

        # Caso não tenha DARF.
        else:
            worksheet.write("G" + str(indice_excel), "❌")

        #---------------------------------------------------------------------------------------------------------------
        # Anexo especial apenas para emails de PARCELAMENTO.
        if (int(tipo_email) == 3) & (codigo_cliente in lista_cod_parcelamentos):

            # Registra no relatório.
            worksheet.write("H" + str(indice_excel), "✔️")

            # Anexa todos os PDFs que constam na lista.
            for indice_parcelamento, codigo_parcelamento in enumerate(lista_cod_parcelamentos):
                if codigo_parcelamento == codigo_cliente:

                    # Constrói o nome do arquivo com o caminho.
                    nome_arquivo_pdf = "Parcelamentos/" + lista_parcelamentos[indice_parcelamento]

                    # Importa e anexa o arquivo.
                    with open(nome_arquivo_pdf, "rb") as arquivo:
                        arquivo_mime = MIMEApplication(arquivo.read())
                        arquivo_mime.add_header("Content-Disposition", "attachment", filename=nome_arquivo_pdf)
                        mensagem.attach(arquivo_mime)

        #---------------------------------------------------------------------------------------------------------------
        # Tenta enviar o email para todos os recipientes.
        try:

            # Envia o email pelo servidor.
            servidor.sendmail(email_origem, email_destino, mensagem.as_string())

            # Marca no relatório que deu CERTO.
            print(f"SUCESSO ao enviar email: {codigo_cliente} | {email_destino}")
            worksheet.write("A" + str(indice_excel), "✔️")

        # Marca no relatório que deu ERRADO.
        except Exception as erro:
            print(f"ERRO ao enviar o email: {codigo_cliente} | {email_destino}", erro)
            worksheet.write("A" + str(indice_excel), "❌")

servidor.quit()
workbook.close()

print("\n\nFIM DO PROGRAMA: Você já pode fechar essa janela :)\n\n")
