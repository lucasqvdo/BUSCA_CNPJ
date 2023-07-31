import requests, json
from openpyxl import Workbook, load_workbook


class fornecedor:
    def __init__(self, nome, fantasia, cnpj, logradouro, numero, bairro, uf, cep, telefone, email):
        self.nome=nome
        self.fantasia=fantasia
        self.cnpj=cnpj
        self.logradouro=logradouro
        self.numero=numero
        self.bairro=bairro
        self.uf=uf
        self.cep=cep
        self.telefone=telefone
        self.email=email

def cadastro_principal():
    
    sheet['C4'].value=fornecedor.nome
    sheet['C3'].value=fornecedor.fantasia
    sheet['C5'].value=fornecedor.cnpj
    sheet['C8'].value=fornecedor.logradouro
    sheet['G8'].value=fornecedor.numero
    sheet['B9'].value=fornecedor.bairro
    sheet['D9'].value=fornecedor.uf
    sheet['G9'].value=fornecedor.cep
    sheet['B10'].value=fornecedor.telefone
    sheet['B11'].value=fornecedor.email



def cadastro_bancario():
    confirma = "n"
    while confirma != "s":

        banco = input("Informe o codigo e nome do banco:\n")
        agencia = input("Informe o numero da agencia:\n")
        conta = input("Informe o número da conta:\n")
        print("Confirmando: \n", "Banco:   ", banco, "\n" ,"Agencia: ", agencia, "\n","Conta:   ", conta)

        confirma = input("Os dados estão corretos? s/n\n")
    
        sheet['B17'].value=banco
        sheet['B18'].value=agencia
        sheet['B19'].value=conta

def salvar_planilha():
    nome = ''.join(filter(str.isalnum, fornecedor.nome)) 
    workbook.save(filename=nome +".xlsx")
    print("Planilha Salva")




while True:
    cnpj=0
    workbook = load_workbook(filename="Cadastro.xlsx")
    sheet = workbook.active


    while len(str(cnpj)) != 14:
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
        
        fornecedor=fornecedor(data['nome'], data['fantasia'], data['cnpj'], data['logradouro'], data['numero'], data['bairro'], data['uf'], data['cep'], data['telefone'], data['email'])
        print("Empresa encontrada: ",fornecedor.nome)
        respostacorreta = False
        while respostacorreta == False:

        
            d_bancarios= input("Deseja adicionar os dados bancários ao cadastro? s/n \n")
            if d_bancarios == "n" or d_bancarios == "N":
                cadastro_principal()
                salvar_planilha()
                respostacorreta = True



            elif d_bancarios == "s" or d_bancarios == "S":
                cadastro_principal()
                cadastro_bancario()
        
                salvar_planilha()
                respostacorreta = True
            
            
            else:
                print("Digite 's' ou 'n':\n")
                
      

    
            

        
        
           

        
        
    
        

    
