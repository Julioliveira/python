import openpyxl
from datetime import datetime

nome = input('Insira o nome do arquivo xlsx: ')
nome += '.xlsx'
wb = openpyxl.load_workbook(filename=nome)
ws = wb.active
datas = []
categoria, subcategoria, origem = '', '', ''
dados = []
for cell in range(2, ws.max_column - 3, 4):
    datas.append({"data": ws.cell(row=1, column=cell).value, "coluna": cell})
for linha in range(1, ws.max_row+1):
    valor_celula = ws.cell(row=linha, column=1).value
    if valor_celula == 'Categ. de despesa':
        origem = 'Despesas'
    elif valor_celula == 'Categ. de rendimento':
        origem = 'Receitas'
    elif str(valor_celula).lower().find('total') > -1:
        continue
    elif valor_celula == None:
        continue
    elif valor_celula[0:2].isnumeric():
        categoria = valor_celula
    elif len(origem) > 0:
            subcategoria = valor_celula
            if subcategoria == 'Todas as outras despesas':
                categoria = 'OUTROS'
            for data in datas:
                real = {
                    'origem': origem,
                    'categoria': categoria,
                    'subcategoria': subcategoria,
                    'data': data['data'],
                    'tipo': 'Real',
                    'valor': ws.cell(row=linha, column=data['coluna']).value
                }
                orcado = {
                    'origem': origem,
                    'categoria': categoria,
                    'subcategoria': subcategoria,
                    'data': data['data'],
                    'tipo': 'Orçado',
                    'valor': ws.cell(row=linha, column=data['coluna']+1).value
                }
                if not real['valor'] == None:
                    dados.append(real)
                if not orcado['valor'] == None:
                    dados.append(orcado)

dados.sort(key=lambda x: (x['origem'], x['data']))
wb.close()
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Real Orçado"
ws['A1'] = 'Origem'
ws['B1'] = 'Categoria'
ws['C1'] = 'Subcategoria'
ws['D1'] = 'Data'
ws['E1'] = 'Tipo'
ws['F1'] = "Valor"
linha = 2
for dado in dados:
    ws.cell(row=linha, column=1, value=dado['origem'])
    ws.cell(row=linha, column=2, value=dado['categoria'])
    ws.cell(row=linha, column=3, value=dado['subcategoria'])
    ws.cell(row=linha, column=4, value=dado['data'])
    ws.cell(row=linha, column=5, value=dado['tipo'])
    ws.cell(row=linha, column=6, value=dado['valor'])
    linha += 1

data = datetime.today()
wb.save(nome)
