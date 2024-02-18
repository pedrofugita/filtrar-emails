from imap_tools import MailBox, AND, NOT
from datetime import date
import openpyxl

# Configurações de login do Google
usuario = "pedrofugita98@gmail.com"
senha = "eptkiczpcetpnyxc"
meu_email = MailBox("imap.gmail.com")
meu_email.login(usuario, senha)

# PALAVRAS-CHAVE PARA BUSCAR
palavras_chave = ["Embraer, contrato, link"]

# PALAVRAS-CHAVE PARA NÃO BUSCAR
palavras_excluir = ["trainee", "bate-papo"]

# EMAILS RECEBIDOS APÓS
data_referencia = date(2023, 9, 1)

# FILTROS
# Filtrar emails com palavras-chave
emails_filtrados = []
for palavra in palavras_chave:
    emails = meu_email.fetch(AND(subject=palavra, date_gte=data_referencia, seen=False))
    emails_filtrados.extend(emails)

# Filtrar emails para exclusão
emails_final = []
for email in emails_filtrados:
    if all(palavra not in email.subject.lower() for palavra in palavras_excluir):
        emails_final.append(email)

# CRIAR PLANILHA EXCEL
# planilha = openpyxl.Workbook()
# planilha_ativa = planilha.active
# planilha_ativa.append(["Data", "Remetente", "Assunto"])
#
# for email in emails_final:
#     data_email = email.date.date()
#     remetente = email.from_
#     assunto = email.subject
#     planilha_ativa.append([data_email, remetente, assunto])
#
# planilha.save("emails_estagio.xlsx")

if len(emails_final) == 1:
    print("\nFoi encontrado apenas 1 email.")
elif len(emails_final) >= 2:
    print(f"\nForam encontrados {len(emails_final)} emails.")
else:
    print("\nNão foi encontrado nenhum email.")
    emails_final = ""

# IMPRIME EMAILS
for i in emails_final:
    print("\nRemetente:", i.from_)
    print("Assunto:", i.subject)
    print()

# FAZER LOGOUT
meu_email.logout()
