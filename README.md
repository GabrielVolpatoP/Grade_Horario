<header>
<h1> GeradorRelatorio.ipynb
<img src="https://img.shields.io/badge/READ%20ME-555555" alt="ReadMe" align="right"  width="q35">
</h1>
</header>

Este projeto tem como objetivo gerar relatórios em PDF a partir de grades de horários extraídas de arquivos PDF. Ele é especialmente útil para organizar e visualizar os horários de disciplinas de diferentes semestres e turnos (matutino e vespertino) de forma automatizada.

## Funcionalidades

- **Leitura de PDF:** Utiliza a biblioteca `tabula` para extrair tabelas de horários de um arquivo PDF.
- **Tratamento dos Dados:** Realiza limpeza e organização dos dados, preenchendo valores nulos, ajustando colunas e separando códigos de disciplinas e salas.
- **Associação de Dados:** Relaciona os códigos das disciplinas com seus respectivos nomes e salas.
- **Geração de Relatório:** Cria arquivos PDF formatados com as tabelas de horários para cada semestre e turno.

## Estrutura do Código

O notebook está organizado em etapas:

1. **Importação de bibliotecas**  
   - tabula, pdfreader, reportlab, entre outras.

2. **Leitura e extração das tabelas do PDF**  
   - O arquivo `horario.pdf` é lido e as tabelas são extraídas para análise.

3. **Tratamento das tabelas**  
   - Remoção de colunas extras, preenchimento de campos vazios, renomeação de colunas.

4. **Separação dos dados**
   - Extrai códigos das disciplinas e das salas, tanto do turno matutino quanto vespertino, para 1º e 2º semestre.

5. **Associação das disciplinas**
   - Junta o código da disciplina com o seu nome e sala, para facilitar a visualização na tabela final.

6. **Geração dos PDFs**
   - Utiliza o `reportlab` para criar arquivos PDF (`Primeiro_Semestre.pdf` e `Segundo_Semestre.pdf`) com as grades de horários formatadas.

## Como usar

1. **Instale as dependências:**
   ```bash
   pip install tabula-py pdfreader reportlab pandas
   ```
   Obs: O Java deve estar instalado para o tabula funcionar.

2. **Coloque o arquivo pdf fornecido pela UEM no portal sisav com o nome de `horario.pdf` na mesma pasta do notebook.**

3. **Execute o notebook:**
   - Siga as células em ordem.
   - Ao final, os PDFs dos horários serão gerados na pasta do projeto.

## Observações

- O código está preparado para tratar arquivos de horários com até 5 páginas/tabelas e fazer a correspondência entre disciplinas e salas.
- Os nomes das disciplinas são truncados para facilitar a visualização.
- O notebook foi desenvolvido em Python 3.

## Possíveis melhorias

- Adicionar tratamento para arquivos com número diferente de tabelas.
- Permitir customização do layout dos PDFs.
- Automatizar a seleção do arquivo de entrada.

## Licença
Este projeto está licenciado sob a [MIT License](LICENSE).

---

## Contato
**Autor:** Gabriel Volpato  
Para suporte ou dúvidas, entre em contato pelo email: volpatocursin@outlook.com 
<img src="https://github.com/GabrielVolpatoP/GabrielVolpatoP/blob/main/imagens/Duck__icon.svg?raw=true" alt="Icon Usuario" align="right"  width="60">
