import pandas as pd

def main():
    print("=== Visualizar tabela Excel ===")

    file = r'C:\Users\fabin\Documents\Planilhas\ControleMensal_v2.xlsx'

    sheet = 'Controle'

    try:
        n = int(input("Quantas últimas linhas deseja ver? [10]: ").strip() or 10)
    except ValueError:
        n = 10

    df = pd.read_excel(file, sheet_name=sheet)

    print(df.tail(n))
    


if __name__ == "__main__":
    main()
