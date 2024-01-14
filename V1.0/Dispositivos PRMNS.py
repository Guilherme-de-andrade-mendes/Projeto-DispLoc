import os
from xhtml2pdf import pisa

class Dispositivo():
    def __init__(self):
        self.nome = ''
        self.status = ''
        self.sistemaOperacional = ''
        self.chaveDeAtivacao = ''
        self.particoes = []

def menu():
    print("="*32 + "Gerenciador de dispositivos locais - Prominas" + "="*32)
    print("1. Nova máquina.\n2. Mostrar todos.\n3. Mostrar uma máquina.\n4. Alterar especificações de uma máquina.\n5. Excluir.\n6. Gerar PDF.\n7. Sair")
    print("="*80)
    opcao = input("Entre com o dígito da função desejada: ")
    return opcao

def existeArquivo(arquivoTxt):
    return os.path.isfile(arquivoTxt)

def leArquivoDeDispositivos(arquivoTxt):
    dispositivos = []
    if existeArquivo(arquivoTxt):
        with open(arquivoTxt, 'r', encoding='utf-8') as arq:
            for linha in arq:
                infos = linha.split(';')
                if len(infos) >= 5:
                    d = Dispositivo()
                    d.nome = infos[0]
                    d.status = infos[1]
                    d.sistemaOperacional = infos[2]
                    d.chaveDeAtivacao = infos[3]
                    d.particoes = infos[4].split(',')
                    dispositivos.append(d)
    else:
        open(arquivoTxt, 'a', encoding='utf-8').close()
    return dispositivos

def gravaDispositivosTxt(dispositivos, arquivoTxt):
    with open(arquivoTxt, 'w', encoding='utf-8') as arq:
        for d in dispositivos:
            particoes = ', '.join(d.particoes)
            arq.write(
                f"{d.nome};{d.status};{d.sistemaOperacional};{d.chaveDeAtivacao};{particoes}\n"
            )

def validaIdentificador(listaDeDispositivos):
    while True:
        y = input("Nome da máquina: ").upper()
        for x in listaDeDispositivos:
            if x.nome == y:
                print("Essa máquina já existe. Tente novamente.")
                break
        else:
            return y

def incluirMaquina(listaDeDispositivos):
    d = Dispositivo()
    print("="*32 + "Cadastrando dispositivo" + "="*32)
    valida = validaIdentificador(listaDeDispositivos)
    d.nome = valida
    exiSta = False
    while not exiSta:
        x = input("Status da máquina (Ativo/Inativo): ").lower().capitalize()
        if x in ["Ativo", "Inativo"]:
            exiSta = True
            d.status = x
    d.sistemaOperacional = input("Sistema Operacional: ").lower().capitalize()
    d.chaveDeAtivacao = input("Chave de ativação: ").upper()
    y = True
    while y:
        particao = input("Espaço de armazenamento da partição: ").upper()
        d.particoes.append(particao)
        pgt = input("Deseja inserir uma nova partição (s/n): ").lower()
        if pgt == "n":
            y = False
    print("Dispositivo inserido com sucesso!")
    listaDeDispositivos.append(d)

def imprimirDados(d):
    print(f"{d.nome} | {d.status} | {d.sistemaOperacional} | {d.chaveDeAtivacao} | [{', '.join(d.particoes)}]")

def mostrarTodos(listaDeDispositivos):
    if len(listaDeDispositivos) > 0:
        print("="*32 + "Listando dispositivo cadastrados" + "="*32 )
        print("Nome | Status | Sistema Operacional | Chave de ativação | Partições\n")
        for dispositivo in listaDeDispositivos:
            imprimirDados(dispositivo)
    else:
        print("Não existem dispositivos cadastrados no sistema.")

def buscarDispositivo(listaDeDispositivos):
    x = input("Nome da máquina: ").upper()
    for y in range(len(listaDeDispositivos)):
        if listaDeDispositivos[y].nome == x:
            return y
    return -1

def imprimirMaquinaEspecifica(listaDeDispositivos):
    print("="*32 + "Mostrando dispositivo específico" + "="*32)
    indDisp = buscarDispositivo(listaDeDispositivos)
    if indDisp >= 0:
        imprimirDados(listaDeDispositivos[indDisp])
    else:
        print("O dispositivo indicado não pode ser encontrado. Verifique se o mesmo consta na base de dados atual.")

def alterarEspecificacao(listaDeDispositivos):
    indDisp = buscarDispositivo(listaDeDispositivos)
    if indDisp >= 0:
        opcao = input("="*32 + "Alterando especificações do dispositivo" + "="*32 + "\n1. Status.\n2. Sistema Operacional.\n3. Chave de ativação.\n4. Partições\nEntre com o dígito da opção que deseja alterar no dispostivo: ")
        if opcao == "1":
            print("Alterando Status do dispositivo\n")
            exiSta = False
            while not exiSta:
                x = input("Status da máquina (Ativo/Inativo): ").lower().capitalize()
                if x in ["Ativo", "Inativo"]:
                    exiSta = True
                    listaDeDispositivos[indDisp].status = x
        elif opcao == "2":
            print("Alterando Sistema Operacional do dispositivo\n")
            listaDeDispositivos[indDisp].sistemaOperacional = input("Sistema Operacional: ").lower().capitalize()
        elif opcao == "3":
            print("Alterando chave de ativação do dispositivo\n")
            listaDeDispositivos[indDisp].chaveDeAtivacao = input("Chave de ativação: ").upper()
        elif opcao == "4":
            print("Alterando partições do dispositivo: ")
            i = True
            x = []
            while i:
                particao = input("Espaço de armazenamento da partição: ").capitalize()
                x.append(f"{particao}")
                pgt = input("Deseja inserir uma nova partição (s/n): ").lower()
                if pgt == "n":
                    i = False
            listaDeDispositivos[indDisp].particoes = x
        else:
            print("Opção inválida. Tente novamente.")
    else:
        print("O dispositivo indicado não pode ser encontrado. Verifique se o mesmo consta na base de dados atual.")

def excluirDispositivo(listaDeDispositivos):
    indDisp = buscarDispositivo(listaDeDispositivos)
    if indDisp >= 0:
        print("="*32 + "Excluindo dispositivo" + "="*32)
        listaDeDispositivos.pop(indDisp)
        print("Dispositivo removido com sucesso!")
    else:
        print("O dispositivo indicado não pode ser encontrado. Verifique se o mesmo consta na base de dados atual.")

def gerarPDF(arquivoTxt):
    print("="*32 + "Gerando PDF" + "="*32)
    pdf_file = "Etiquetas para dispositivos.pdf"

    # Criar um documento HTML
    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <style>
            body {{
                font-family: 'Arial Narrow', Arial, sans-serif;
            }}
            .container {{
                display: flex;
                flex-wrap: wrap;
            }}
            table {{
                width: 100%;
            }}
            th, td {{
                border: 1px solid black;
                text-align: center;
                padding: 8px;
                font-size: medium;
            }}
            th {{
                background-color: #f2f2f2;
                text-align: center;
            }}
            tr > td{{
                font-size: 9pt;
            }}
        </style>
    </head>
    <body>
        <div class="container">
    """

    # Adicionar informações de cada máquina ao HTML
    with open(arquivoTxt, 'r', encoding='utf-8') as arq:
        for linha in arq:
            infos = linha.split(';')
            if len(infos) >= 5:
                html_content += f"""
                <div class="column">
                    <table>
                        <tr>
                            <th><h1>Nome do dispositivo</h1></th>
                            <th><h1>Sistema Operacional</h1></th>
                            <th><h1>Chave de Ativação</h1></th>
                        </tr>
                        <tr>
                            <td>{infos[0]}</td>
                            <td>{infos[2]}</td>
                            <td>{infos[3]}</td>
                        </tr>
                    </table>
                </div>
                """
                    
    # Fechar o HTML
    html_content += """
        </div>
    </body>
    </html>
    """

    # Criar o PDF a partir do HTML
    with open(pdf_file, 'w+b') as pdf:
        pisa.CreatePDF(html_content, dest=pdf)

    print(f"PDF gerado com sucesso em {pdf_file}")


def main():
    arquivoDispositivos = 'Dispositivos Prominas.txt'
    dispositivos = leArquivoDeDispositivos(arquivoDispositivos)
    opcao = menu()
    while opcao != "7":
        if opcao == "1":
            incluirMaquina(dispositivos)
        elif opcao == "2":
            mostrarTodos(dispositivos)
        elif opcao == "3":
            imprimirMaquinaEspecifica(dispositivos)
        elif opcao == "4":
            alterarEspecificacao(dispositivos)
        elif opcao == "5":
            excluirDispositivo(dispositivos)
        elif opcao == "6":
            gerarPDF(arquivoDispositivos)
        else:
            print("Opção inválida. Tente novamente com uma das opções disponíveis no menu.")
        gravaDispositivosTxt(dispositivos, arquivoDispositivos)
        opcao = menu()
    gravaDispositivosTxt(dispositivos, arquivoDispositivos)
    print("Desligando...")

if __name__ == "__main__":
    main()