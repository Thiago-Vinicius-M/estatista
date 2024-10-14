import pandas as pd

# Caminho do arquivo
arquivo_excel = "C:/Users/Thiago Dias/Desktop/ProjetoEstatistica/Projeto Estatística.xlsx"

# Nomeando as páginas da planilha
planilhas = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Cria um DataFrame vazio para armazenar todos os dados
all_data = pd.DataFrame()
p
# Iterar sobre a planilha e concatenar os dados
for planilha in planilhas:
    df = pd.read_excel(arquivo_excel, sheet_name=planilha)
    all_data = pd.concat([all_data, df], ignore_index=True)

# Converter colunas numéricas para o tipo float
all_data['Valor da venda'] = pd.to_numeric(all_data['Valor da venda'])

# Função para eleger o melhor vendedor de acordo com o critério selecionado
def melhor_vendedor(df, criterio='valor_total'):
    """
    Argumentos:
        df: DataFrame com os dados de vendas.
        criterio: Critério de análise (valor_total, num_vendas, media_valor).

    Returns:
        Uma string com o nome do vendedor
    """
    if criterio == 'valor_total':
        return df.groupby('Vendedor')['Valor da venda'].sum().idxmax()
    elif criterio == 'num_vendas':
        return df.groupby('Vendedor')['Vendedor'].count().idxmax()
    elif criterio == 'media_valor':
        return df.groupby('Vendedor')['Valor da venda'].mean().idxmax()
    else:
        raise ValueError("Critério inválido. Opções válidas: 'valor_total', 'num_vendas', 'media_valor'")
    
# Menu interativo
def main():
    while True:
        print('-----------------------------------------------------')
        print("Escolha uma opção")
        print("1. Analisar todas as vendas")
        print("2. Analisar vendas por mês")
        print("3. Sair")
        opcao = input("Digite o número da opção: ")
        print('-----------------------------------------------------')

        # Se a opção for todas as vendas, puxar o DataFrame com todos os dados concatenados
        if opcao == '1':
            df = all_data

        # Se a opção for Analisar as vendas do mês, o usuário indica o mês
        elif opcao == '2':
            while True:
                try:
                    print('Selecione o mês')
                    print('Exemplo: 1 - Janeiro, 2 Fevereiro, 3 - Março e etc...')
                    mes = int(input("Digite o número correspondente ao mês: "))
                    print('-----------------------------------------------------')
                    if 1 <= mes <= 12:
                        break
                    else:
                        print("Número de mês inválido. Digite um número entre 1 e 12.")
                        print('-----------------------------------------------------')
                except ValueError:
                    print("Entrada inválida. Digite um número inteiro.")
                    print('-----------------------------------------------------')

            # Carregar o DataFrame para o mês selecionado
            nome_planilha = planilhas[mes - 1]
            df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)

        # A opção 3-Sair encerra o loop
        elif opcao == '3':
            break
        
        else:
            print("Opção inválida.")
            continue

        # Se a opção escolhida for 1 ou 2, o sistema puxa os dados do DataFrame        
        if opcao in ['1', '2']:
            criterio = obter_criterio()
            melhor = melhor_vendedor(df, criterio)

            # Agrupa as colunas 'Vendedor' e 'Valor da venda' e soma, caso o critério seja Valor Total)
            if criterio == 'valor_total':
                valor_total = df.groupby('Vendedor')['Valor da venda'].sum()[melhor]
                print('-----------------------------------------------------')
                print(f"O melhor vendedor em termos de valor total é {melhor}!") 
                print(f"Com R${valor_total:.2f}.")
            
            # Agrupa a 'Vendedor' e conta a quantidade de linhas, já que cada venda é representada por uma linha
            elif criterio == 'num_vendas':
                num_vendas = df.groupby('Vendedor')['Vendedor'].count()[melhor]
                print('-----------------------------------------------------')
                print(f"O melhor vendedor em número de vendas é {melhor}!")
                print(f"Com {num_vendas} vendas.")
            
            # Agrupa as colunas 'Vendedor' e Valor da venda e tira a média
            elif criterio == 'media_valor':
                media_valor = df.groupby('Vendedor')['Valor da venda'].mean()[melhor]
                print('-----------------------------------------------------')
                print(f"O melhor vendedor em média por venda é {melhor}!")
                print(f"Com R${media_valor:.2f} por venda.")

# Função para obter do usuário, o critério
def obter_criterio():
    while True:
        print('Escolha o critério')
        print('1 - Valor acumulado de vendas')
        print('2 - Maior número de vendas')
        print('3 - Média por venda')
        try:
            criterio = int(input('Digite o número da opção: '))

            # Verifica se o número digitado está entre 1 e 3
            if 1 <= criterio <= 3:

                # Retorna o critério
                if criterio == 1:
                    return 'valor_total'
                elif criterio == 2:
                    return 'num_vendas'
                else:
                    return 'media_valor'
            else:
                print("Opção inválida. Digite 1, 2 ou 3.")
        except ValueError:
            print("Entrada inválida. Digite um número inteiro.")

if __name__ == "__main__":
    main()