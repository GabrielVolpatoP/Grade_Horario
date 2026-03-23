import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import warnings
import os
import sys

warnings.filterwarnings("ignore") 

print("==================================================")
print("      GERADOR DE RELATÓRIO DE HORÁRIOS")
print("==================================================")

#arquivo_pdf = input("Digite o nome do arquivo PDF (ex: horario.pdf): ")

# Verifica se o arquivo existe antes de continuar
if not os.path.exists("SAV - Horário de Aulas.pdf"):
    print(f"\nERRO: O arquivo '{arquivo_pdf}' não foi encontrado na pasta atual.")
    input("Pressione ENTER para sair...")
    sys.exit()


print("\nLendo o PDF e extraindo tabelas...")
lista_tabelas = []


with pdfplumber.open('SAV - Horário de Aulas.pdf') as pdf:
    for page in pdf.pages:
        tabelas_extraidas = page.extract_tables()
        for tabela in tabelas_extraidas:
            # Converte a lista do pdfplumber para um DataFrame do pandas
            # Tabela[0] costuma ser o cabeçalho, tabela[1:] são os dados
            df = pd.DataFrame(tabela[1:], columns=tabela[0])
            lista_tabelas.append(df)

print(f"Total de tabelas carregadas: {len(lista_tabelas)}")

if len(lista_tabelas) < 5:
    print("ERRO: O PDF não possui o número esperado de tabelas.")
    input("Pressione ENTER para sair...")
    sys.exit()

# 2. TRATAMENTO: 1º Semestre

# Substituindo vazios. O pdfplumber pode retornar None ou strings vazias
Grade_M1 = lista_tabelas[0].replace("", pd.NA).fillna(value="0 - \n")
Grade_V1 = lista_tabelas[1].replace("", pd.NA).fillna(value="0 - \n")

# Remove colunas se existirem (o pdfplumber pode não criar 'Unnamed: 0', então verificamos antes)
if 'Unnamed: 0' in Grade_M1.columns: Grade_M1 = Grade_M1.drop('Unnamed: 0', axis=1)
if 'Unnamed: 0' in Grade_V1.columns: Grade_V1 = Grade_V1.drop('Unnamed: 0', axis=1)
if 'Hr./Aula' in Grade_M1.columns: Grade_M1 = Grade_M1.drop('Hr./Aula', axis=1)
if 'Hr./Aula' in Grade_V1.columns: Grade_V1 = Grade_V1.drop('Hr./Aula', axis=1)
if 'Sábado' in Grade_M1.columns: Grade_M1 = Grade_M1.drop('Sábado', axis=1)
if 'Sábado' in Grade_V1.columns: Grade_V1 = Grade_V1.drop('Sábado', axis=1)


Grade_M1.columns = ["Horario", 1, 2, 3, 4, 5]
Grade_V1.columns = ["Horario", 1, 2, 3, 4, 5]


for h in range(0, 5):
    Grade_M1.iat[h, 0] = str(Grade_M1.iat[h, 0])[:5] + "\n" + str(Grade_M1.iat[h, 0])[6:]
for h in range(0, 5):
    Grade_V1.iat[h, 0] = str(Grade_V1.iat[h, 0])[:5] + "\n" + str(Grade_V1.iat[h, 0])[6:]

Cod_Disc_M1 = [[], [], [], [], []]
Cod_Sala_M1 = [[], [], [], [], []]
for i in range(1, 6):
    Cod_Disc_M1[i-1] = Grade_M1[i].astype(str).str.split(" - ", expand=True)[0]
    # pdfplumber usa \n em vez de \r para quebras de linha em células
    Cod_Sala_M1[i-1] = Grade_M1[i].astype(str).str.split("\n", expand=True)[1]

Cod_Disc_V1 = [[], [], [], [], []]
Cod_Sala_V1 = [[], [], [], [], []]
for i in range(1, 6):
    Cod_Disc_V1[i-1] = Grade_V1[i].astype(str).str.split(" - ", expand=True)[0]
    Cod_Sala_V1[i-1] = Grade_V1[i].astype(str).str.split("\n", expand=True)[1]

# 3. TRATAMENTO: 2º Semestre

Grade_M2 = lista_tabelas[2].replace("", pd.NA).fillna(value="0 - \n")
Grade_V2 = lista_tabelas[3].replace("", pd.NA).fillna(value="0 - \n")

if 'Unnamed: 0' in Grade_M2.columns: Grade_M2 = Grade_M2.drop('Unnamed: 0', axis=1)
if 'Unnamed: 0' in Grade_V2.columns: Grade_V2 = Grade_V2.drop('Unnamed: 0', axis=1)
if 'Hr./Aula' in Grade_M2.columns: Grade_M2 = Grade_M2.drop('Hr./Aula', axis=1)
if 'Hr./Aula' in Grade_V2.columns: Grade_V2 = Grade_V2.drop('Hr./Aula', axis=1)
if 'Sábado' in Grade_M2.columns: Grade_M2 = Grade_M2.drop('Sábado', axis=1)
if 'Sábado' in Grade_V2.columns: Grade_V2 = Grade_V2.drop('Sábado', axis=1)

Grade_M2.columns = ["Horario", 1, 2, 3, 4, 5]
Grade_V2.columns = ["Horario", 1, 2, 3, 4, 5]

for h in range(0, 5):
    Grade_M2.iat[h, 0] = str(Grade_M2.iat[h, 0])[:5] + "\n" + str(Grade_M2.iat[h, 0])[6:]
for h in range(0, 5):
    Grade_V2.iat[h, 0] = str(Grade_V2.iat[h, 0])[:5] + "\n" + str(Grade_V2.iat[h, 0])[6:]

Cod_Disc_M2 = [[], [], [], [], []]
Cod_Sala_M2 = [[], [], [], [], []]
for i in range(1, 6):
    Cod_Disc_M2[i-1] = Grade_M2[i].astype(str).str.split(" - ", expand=True)[0]
    Cod_Sala_M2[i-1] = Grade_M2[i].astype(str).str.split("\n", expand=True)[1]

Cod_Disc_V2 = [[], [], [], [], []]
Cod_Sala_V2 = [[], [], [], [], []]
for i in range(1, 6):
    Cod_Disc_V2[i-1] = Grade_V2[i].astype(str).str.split(" - ", expand=True)[0]
    Cod_Sala_V2[i-1] = Grade_V2[i].astype(str).str.split("\n", expand=True)[1]

# 4. LISTA DE DISCIPLINAS

diciplinas = lista_tabelas[len(lista_tabelas)-1]
coluna_disciplina = diciplinas['Disciplina']
diciplinas[diciplinas.columns[0]] = diciplinas[diciplinas.columns[0]].fillna(value=int(1))
dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]

print("Gerando arquivos PDF de saída...")


# 5. GERAÇÃO PDF: 1º Semestre

Grade_M1.columns = ['Horário',] + dias
Grade_V1.columns = ['Horário',] + dias

# Preenchendo Grade_M1 (Matutino)
for h in range(0, 5):
    for j in range(0, 5):
        # Limpa espaços em branco por segurança
        cod_grade_m1 = str(Cod_Disc_M1[h][j]).strip()
        
        if cod_grade_m1 == "0" or cod_grade_m1 == "nan":
            Grade_M1.at[j, dias[h]] = " "
        else:
            for x in range(0, len(diciplinas)):
                # Tratamento de segurança: remove o '.0' caso o pandas tenha lido como float
                cod_tabela = str(diciplinas.iat[x, 0]).replace(".0", "").strip()
                
                if cod_tabela == cod_grade_m1:
                    nome_disc = str(diciplinas.iat[x, 2])
                    sala_m1 = str(Cod_Sala_M1[h][j]).strip() # Corrigido para Cod_Sala_M1
                    
                    # Usa o .at[] que é a forma mais recomendada/segura do Pandas para editar um valor na célula
                    Grade_M1.at[j, dias[h]] = nome_disc[:22] + "\n" + nome_disc[22:] + " - " + sala_m1
                    break

# Preenchendo Grade_V1 (Vespertino)
for h in range(0, 5):
    for j in range(0, 5):
        cod_grade_v1 = str(Cod_Disc_V1[h][j]).strip()
        
        if cod_grade_v1 == "0" or cod_grade_v1 == "nan":
            Grade_V1.at[j, dias[h]] = " "
        else:
            for x in range(0, len(diciplinas)):
                cod_tabela = str(diciplinas.iat[x, 0]).replace(".0", "").strip()
                
                if cod_tabela == cod_grade_v1:
                    nome_disc = str(diciplinas.iat[x, 2])
                    sala_v1 = str(Cod_Sala_V1[h][j]).strip()
                    
                    Grade_V1.at[j, dias[h]] = nome_disc[:22] + "\n" + nome_disc[22:] + " - " + sala_v1
                    break

titulo_pdf1 = list(Grade_M1.columns)
teste1 = [titulo_pdf1]

for j in range(0, len(titulo_pdf1)-1):
    p0 = [Grade_M1[i][j] for i in titulo_pdf1]
    teste1.append(p0)

for j in range(0, len(titulo_pdf1)-1):
    p1 = [Grade_V1[i][j] for i in titulo_pdf1]
    teste1.append(p1)

def create_pdf(file_name, data):
    page_size = landscape(letter)
    doc = SimpleDocTemplate(file_name, pagesize=page_size)
    elements = []
    def_tab = 140
    table_dias = Table(data, colWidths=[48, def_tab, def_tab, def_tab, def_tab, def_tab])
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (1, 1), (-1, -1), 8),
        ('FONTSIZE', (0, 1), (0, -1), 9),
    ])
    table_dias.setStyle(style)
    elements.append(table_dias)
    doc.build(elements)

create_pdf("Primeiro_Semestre.pdf", teste1)


# 6. GERAÇÃO PDF: 2º Semestre


Grade_M2.columns = ['Horário',] + dias
Grade_V2.columns = ['Horário',] + dias

# Preenchendo Grade_M2 (Matutino)
for h in range(0, 5):
    for j in range(0, 5):
        # Limpa espaços em branco por segurança
        cod_grade_M2 = str(Cod_Disc_M2[h][j]).strip()
        
        if cod_grade_M2 == "0" or cod_grade_M2 == "nan":
            Grade_M2.at[j, dias[h]] = " "
        else:
            for x in range(0, len(diciplinas)):
                # Tratamento de segurança: remove o '.0' caso o pandas tenha lido como float
                cod_tabela = str(diciplinas.iat[x, 0]).replace(".0", "").strip()
                
                if cod_tabela == cod_grade_M2:
                    nome_disc = str(diciplinas.iat[x, 2])
                    sala_M2 = str(Cod_Sala_M2[h][j]).strip() # Corrigido para Cod_Sala_M2
                    
                    # Usa o .at[] que é a forma mais recomendada/segura do Pandas para editar um valor na célula
                    Grade_M2.at[j, dias[h]] = nome_disc[:22] + "\n" + nome_disc[22:] + " - " + sala_M2
                    break

# Preenchendo Grade_V2 (Vespertino)
for h in range(0, 5):
    for j in range(0, 5):
        cod_grade_V2 = str(Cod_Disc_V2[h][j]).strip()
        
        if cod_grade_V2 == "0" or cod_grade_V2 == "nan":
            Grade_V2.at[j, dias[h]] = " "
        else:
            for x in range(0, len(diciplinas)):
                cod_tabela = str(diciplinas.iat[x, 0]).replace(".0", "").strip()
                
                if cod_tabela == cod_grade_V2:
                    nome_disc = str(diciplinas.iat[x, 2])
                    sala_V2 = str(Cod_Sala_V2[h][j]).strip()
                    
                    Grade_V2.at[j, dias[h]] = nome_disc[:22] + "\n" + nome_disc[22:] + " - " + sala_V2
                    break

titulo_pdf2 = list(Grade_M2.columns)
teste2 = [titulo_pdf2]

for j in range(0, len(titulo_pdf2)-1):
    p0 = [Grade_M2[i][j] for i in titulo_pdf2]
    teste2.append(p0)

for j in range(0, len(titulo_pdf2)-1):
    p1 = [Grade_V2[i][j] for i in titulo_pdf2]
    teste2.append(p1)

create_pdf("Segundo_Semestre.pdf", teste2)

print("\nSucesso! Arquivos 'Primeiro_Semestre.pdf' e 'Segundo_Semestre.pdf' gerados.")
input("Pressione ENTER para finalizar...")