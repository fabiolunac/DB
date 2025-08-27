from openpyxl import load_workbook
from datetime import datetime, date

ANO_PADRAO = 2025



def ask(prompt, default=None, required=False):
    while True:
        txt = input(f"{prompt}{' ['+default+']' if default else ''}: ").strip()
        if not txt and default is not None:
            return default
        if required and not txt:
            print("Campo obrigatório. Tente novamente.")
            continue
        return txt
    
def parse_number(s):
    if s == "" or s is None:
        return None
    s = s.replace(".", "").replace(",", ".")  
    return float(s)

def parse_date_custom(s):
    """
    Converte string em uma data:
    - 'dd/mm/aaaa' ou 'aaaa-mm-dd' (normal)
    - 'ddmm' → assume ano fixo (ex: 2408 -> 24/08/2025)
    """
    # formato ddmm -> assume ano 2025
    if len(s) == 4 and s.isdigit():
        dia = int(s[:2])
        mes = int(s[2:])
        return date(ANO_PADRAO, mes, dia)

    # formatos normais
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    raise ValueError("Data inválida. Use ddmm, dd/mm/aaaa ou aaaa-mm-dd.")

def add_row(file, shett_name, data, tipo, valor_euro, valor_chf, local, descricao, categoria):
    wb = load_workbook(file)
    ws = wb[shett_name]

    new_line = [
        None, 
        data, 
        tipo,
        valor_euro, 
        valor_chf,
        local, 
        descricao, 
        categoria
    ]

    ws.append(new_line)
    wb.save(file)


def main():

    file = r'C:\Users\fabin\Documents\Planilhas\ControleMensal_v2.xlsx'
    sheet = 'Controle'

    print('===========================================')
    print('= Script para adicionar linhas na tabela ==')
    print('===========================================')

    while True:
        # --------- Date --------- 
        hoje = datetime.today().strftime("%d%m")
        while True:
            try:
                data_str = ask("Data (ddmm)", default=hoje, required=True)
                data_val = parse_date_custom(data_str)
                break
            except ValueError as e:
                print(e)

        
        # --------- Tipo --------- 
        tipo = ask("Tipo (ex: In/Out) [Out]", default= "Out", required=True)

        # --------- Valor Euro --------- 
        veuro_str = ask("Valor em Euro [0]", default=0)
        veuro = parse_number(veuro_str) if veuro_str else veuro_str

        # --------- Valor CHF --------- 
        while True:
            try:
                vchf_str = ask("Valor (CHF)", required=True)
                valor_chf = parse_number(vchf_str)
                break
            except ValueError:
                print("Número inválido. Use 23,55 ou 23.55.")

        # --------- Local ---------
        local     = ask("Local", required=True)

        # --------- Descrição --------- 
        descricao = ask("Descrição", required=True)

        # --------- Categoria --------- 
        categoria = ask("Categoria", required=True)

        add_row(
            file= file,
            shett_name= sheet,
            data=data_val,
            tipo=tipo,
            valor_euro=veuro, 
            valor_chf=valor_chf, 
            local=local, 
            descricao=descricao, 
            categoria=categoria
        )

        print('Linha adicionada com sucesso!\n')

        cont = input("Deseja adicionar outra linha? (s/n): ").strip().lower()
        if cont != "s":
            print("Saindo...")
            break

        print('\n')

if __name__ == "__main__":
    main()