import requests, json
from openpyxl import Workbook, load_workbook
while True:
    cnpj=input("Digite o CNPJ que será consultado:\n")

    url = "https://receitaws.com.br/v1/cnpj/" + cnpj
    params = {
        "param1": "value1",
        "param2": "value2"
    }
    headers = {
        "Authorization": "Bearer YOUR_ACCESS_TOKEN"
    }

    response = requests.get(url, params=params, headers=headers)


    if response.status_code == 200:
        data = response.json()
        print(data['nome'])

        workbook = load_workbook(filename="Cadastro.xlsx")
        sheet = workbook.active
        sheet['C7'].value=data['nome']
        sheet['C8'].value=data['cnpj']
        sheet['B11'].value=data['logradouro']
        sheet['H11'].value=data['numero']
        sheet['B12'].value=data['bairro']
        sheet['D12'].value=data['uf']
        sheet['H12'].value=data['cep']
        sheet['B13'].value=data['telefone']
        sheet['B14'].value=data['email']

        workbook.save(filename=data['nome']+".xlsx")
    else:
        print("Error: ", response.status_code)
