import tabula
import pdfreader
from pdfreader import PDFDocument, SimplePDFViewer
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

import warnings
warnings.filterwarnings("ignore") 

# Carregando o arquivo PDF
lista_tabelas = tabula.read_pdf("horario.pdf", pages = 'all')

# Analise da quantas páginas que foram carregadas
print(len(lista_tabelas))

# Criação e preenchimento das TABELAS

##### 1º Semestre

# Obtendo somente a planilha 1 e transcrevendo os valores vazios com "0 - \r"
Grade_M1 = lista_tabelas[0].fillna(value="0 - \r")

# Obtendo somente a planilha 2 e transcrevendo os valores vazios com "0 - \r"
Grade_V1 = lista_tabelas[1].fillna(value="0 - \r")

# 0 - \r" serão usados para divisões das colunas futuramente
Grade_M1 = Grade_M1.drop('Unnamed: 0', axis = 1)
Grade_V1 = Grade_V1.drop('Unnamed: 0', axis = 1)
Grade_M1 = Grade_M1.drop('Hr./Aula', axis = 1)
Grade_V1 = Grade_V1.drop('Hr./Aula', axis = 1)

# Removendo as colunas sobresalentes das planilhas
Grade_M1.columns = ["Horario",1,2,3,4,5]
Grade_V1.columns = ["Horario",1,2,3,4,5]

##### 2º Semestre

# Obtendo somente a planilha 1 e transcrevendo os valores vazios com "0 - \r"
Grade_M2 = lista_tabelas[2].fillna(value="0 - \r")

# Obtendo somente a planilha 2 e transcrevendo os valores vazios com "0 - \r"
Grade_V2 = lista_tabelas[3].fillna(value="0 - \r")

# 0 - \r" serão usados para divisões das colunas futuramente
Grade_M2 = Grade_M2.drop('Unnamed: 0', axis = 1)
Grade_V2 = Grade_V2.drop('Unnamed: 0', axis = 1)
Grade_M2 = Grade_M2.drop('Hr./Aula', axis = 1)
Grade_V2 = Grade_V2.drop('Hr./Aula', axis = 1)

# Removendo as colunas sobresalentes das planilhas
Grade_M2.columns = ["Horario",1,2,3,4,5]
Grade_V2.columns = ["Horario",1,2,3,4,5]

# TRATAMENTO DAS COLUNAS

##### 1º Semestre

for h in range (0,5):
    Grade_M1.iat[h,0] = Grade_M1.iat[h,0][:5] + "\n" + Grade_M1.iat[h,0][6:]

for h in range (0,5):
    Grade_V1.iat[h,0] = Grade_V1.iat[h,0][:5] + "\n" + Grade_V1.iat[h,0][6:]

##### 2º Semestre

for h in range (0,5):
    Grade_M2.iat[h,0] = Grade_M2.iat[h,0][:5] + "\n" + Grade_M2.iat[h,0][6:]

for h in range (0,5):
    Grade_V2.iat[h,0] = Grade_V2.iat[h,0][:5] + "\n" + Grade_V2.iat[h,0][6:]

### COLETA DE DADOS DAS DICIPLINAS

##### 1º Semestre

# Matutino
Cod_Disc_M1 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Disc_M1[i-1] = Grade_M1[i].str.split(" - ", expand = True)[0]


# Vespertino
Cod_Disc_V1 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Disc_V1[i-1] = Grade_V1[i].str.split(" - ", expand = True)[0]


##### 2º Semestre

# Matutino
Cod_Disc_M2 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Disc_M2[i-1] = Grade_M2[i].str.split(" - ", expand = True)[0]


# Vespertino
Cod_Disc_V2 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Disc_V2[i-1] = Grade_V2[i].str.split(" - ", expand = True)[0]


### COLETA DE DADOS DAS SALAS

##### 1º Semestre

# Matutino
Cod_Sala_M1 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Sala_M1[i-1] = Grade_M1[i].str.split("\r", expand = True)[1]


# Vespertino
Cod_Sala_V1 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Sala_V1[i-1] = Grade_V1[i].str.split("\r", expand = True)[1]


##### 2º Semestre

# Matutino
Cod_Sala_M2 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Sala_M2[i-1] = Grade_M2[i].str.split("\r", expand = True)[1]


# Vespertino
Cod_Sala_V2 = [[],[],[],[],[]]
for i in range(1,6):
    Cod_Sala_V2[i-1] = Grade_V2[i].str.split("\r", expand = True)[1]


*FIM DO TRATAMENTO DE COLUNAS*

# Coleta *(Lista)* das diciplinas

#Localizando a tabela com as diciplinas (por padrão a ultima planilha)
diciplinas = lista_tabelas[len(lista_tabelas)-1]

coluna_diciplina = diciplinas.columns[2]

#Preenchendo os valores NULOS (NaN) da coluna de codicos das diciplinas por um valor (qualquer n nulo)
diciplinas[diciplinas.columns[0]] = diciplinas[diciplinas.columns[0]].fillna(value=int(1))

# Tratamento e associação das tabelas com a lista de diciplinas

##### 1º Semestre

dias = ["Segunda","Terça","Quarta","Quinta","Sexta"]

Grade_M1.columns = ['Horário',]+dias
Grade_V1.columns = ['Horário',]+dias

# Associando os codicos da grade dos dias das semanas com os codicos da tabela de diciplinas
for h in range(0,5):
    for j in range(0,5):
        for x in range(0,len(diciplinas)):
            # Caso o dia esteja sem aulas o horario permanece vazio
            if int(Cod_Disc_M1[h][j]) == 0:
                Grade_M1[dias[h]][j] = " "
                break
            else:
                # Reescrevimento da grade com os nomes das diciplinas
                if str(diciplinas.iat[x, 0]) == str(Cod_Disc_M1[h][j]):
                    Grade_M1[dias[h]][j] = diciplinas.iat[x, 2][:22] + "\n" + diciplinas.iat[x, 2][22:] + " - " + Cod_Disc_M1[h][j]
                    break

# Associando os codicos da grade dos dias das semanas com os codicos da tabela de diciplinas
for h in range(0,5):
    for j in range(0,5):
        for x in range(0,len(diciplinas)):
            # Caso o dia esteja sem aulas o horario permanece vazio
            if int(Cod_Disc_V1[h][j]) == 0:
                Grade_V1[dias[h]][j] = " "
                break
            else:
                # Reescrevimento da grade com os nomes das diciplinas
                if str(diciplinas.iat[x, 0]) == str(Cod_Disc_V1[h][j]):
                    Grade_V1[dias[h]][j] = diciplinas.iat[x, 2][:22] + "\n" + diciplinas.iat[x, 2][22:] + " - " + Cod_Sala_V1[h][j]
                    break

titulo_pdf = []
for i in Grade_M1.columns:
    titulo_pdf += i,


teste = [titulo_pdf]

for j in range(0,len(titulo_pdf)-1):
    p0 = []
    for i in titulo_pdf:
        p0 += (Grade_M1[i][j]),
    teste += p0,

for j in range(0,len(titulo_pdf)-1):
    p1 = []
    for i in titulo_pdf:
        p1 += (Grade_V1[i][j]),
    teste += p1,

def create_pdf(file_name):
    
    page_size = landscape(letter)
    
    # Cria um documento
    doc = SimpleDocTemplate(file_name, pagesize=page_size)
    
    # Lista de elementos a serem adicionados ao documento
    elements = []

    # Define os dados da tabela
    horario_pdf = [teste]
    
    # Cria a tabela
    def_tab = 140
    table_dias = Table(teste, colWidths = [48, def_tab, def_tab, def_tab, def_tab, def_tab])
    
    # Define o estilo da tabela
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
    
    # Adiciona a tabela à lista de elementos
    #elements.append(table)
    elements.append(table_dias)
    
    # Constrói o documento
    doc.build(elements)

if __name__ == "__main__":
    create_pdf("Primeiro_Semestre.pdf")


##### 2º Semestre

dias = ["Segunda","Terça","Quarta","Quinta","Sexta"]

Grade_M2.columns = ['Horário',]+dias
Grade_V2.columns = ['Horário',]+dias

# Associando os codicos da grade dos dias das semanas com os codicos da tabela de diciplinas
for h in range(0,5):
    for j in range(0,5):
        for x in range(0,len(diciplinas)):
            # Caso o dia esteja sem aulas o horario permanece vazio
            if int(Cod_Disc_M2[h][j]) == 0:
                Grade_M2[dias[h]][j] = " "
                break
            else:
                # Reescrevimento da grade com os nomes das diciplinas
                if str(diciplinas.iat[x, 0]) == str(Cod_Disc_M2[h][j]):
                    Grade_M2[dias[h]][j] = diciplinas.iat[x, 2][:22] + "\n" + diciplinas.iat[x, 2][22:] + " - " + Cod_Disc_M2[h][j]
                    break

# Associando os codicos da grade dos dias das semanas com os codicos da tabela de diciplinas
for h in range(0,5):
    for j in range(0,5):
        for x in range(0,len(diciplinas)):
            # Caso o dia esteja sem aulas o horario permanece vazio
            if int(Cod_Disc_V2[h][j]) == 0:
                Grade_V2[dias[h]][j] = " "
                break
            else:
                # Reescrevimento da grade com os nomes das diciplinas
                if str(diciplinas.iat[x, 0]) == str(Cod_Disc_V2[h][j]):
                    Grade_V2[dias[h]][j] = diciplinas.iat[x, 2][:22] + "\n" + diciplinas.iat[x, 2][22:] + " - " + Cod_Sala_V2[h][j]
                    break

titulo_pdf = []
for i in Grade_M2.columns:
    titulo_pdf += i,


teste = [titulo_pdf]

for j in range(0,len(titulo_pdf)-1):
    p0 = []
    for i in titulo_pdf:
        p0 += (Grade_M2[i][j]),
    teste += p0,

for j in range(0,len(titulo_pdf)-1):
    p1 = []
    for i in titulo_pdf:
        p1 += (Grade_V2[i][j]),
    teste += p1,

def create_pdf(file_name):
    
    page_size = landscape(letter)
    
    # Cria um documento
    doc = SimpleDocTemplate(file_name, pagesize=page_size)
    
    # Lista de elementos a serem adicionados ao documento
    elements = []

    # Define os dados da tabela
    horario_pdf = [teste]
    
    # Cria a tabela
    def_tab = 140
    table_dias = Table(teste, colWidths = [48, def_tab, def_tab, def_tab, def_tab, def_tab])
    
    # Define o estilo da tabela
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
    
    # Adiciona a tabela à lista de elementos
    #elements.append(table)
    elements.append(table_dias)
    
    # Constrói o documento
    doc.build(elements)

if __name__ == "__main__":
    create_pdf("Segundo_Semestre.pdf")
