import pandas as pd
import numpy as np
import chardet


def carregar_planilhas():
    print("Carregando arquivo Excel...")

    margem_df = pd.read_excel(
        r"C:\Users\DELL\Downloads\260520_MRG - wapp.xlsx",
        sheet_name="Base (3,5%)",
        header=8
    )

    print("Carregando arquivo CSV...")

    csv_path = r"Y:\hor\excel\20230101.csv"

    # detectar encoding (amostra ao invés de arquivo inteiro)
    with open(csv_path, 'rb') as f:
        raw_data = f.read(200000)
        encoding = chardet.detect(raw_data)['encoding']

    encodings = list(dict.fromkeys([
        encoding, 'cp1252', 'latin-1', 'iso-8859-1', 'utf-8'
    ]))

    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                first_line = f.readline()
                sep = ';' if first_line.count(';') > first_line.count(',') else ','

            csv_df = pd.read_csv(
                csv_path,
                encoding=enc,
                sep=sep,
                engine='python',
                on_bad_lines='skip',
                decimal=',',
                thousands='.',
                dtype={'HISTORICO': str}
            )

            print(f"CSV carregado com encoding: {enc}")
            break

        except Exception:
            continue
    else:
        csv_df = pd.read_csv(csv_path, sep=sep, engine='python')

    return margem_df, csv_df


def limpar_dados(margem_df, csv_df):

    print("\nPreparando dados...")

    csv_df = csv_df.rename(columns={
        'ROMANEIO': 'OS',
        'NOTA FISCAL': 'NF',
        'PRODUTO': 'CODPRODUTO',
        'PESO': 'PESO_CSV',
        'UNITARIO': 'PRECO_CSV',
        'HISTORICO': 'HISTORICO_CSV'
    })

    margem_df = margem_df[
        ['OS', 'NF-E', 'CODPRODUTO', 'QTDE AJUSTADA', 'Preço Venda ', 'CF']
    ].copy()

    # conversões
    for col in ['QTDE AJUSTADA', 'Preço Venda ']:
        margem_df[col] = pd.to_numeric(margem_df[col], errors='coerce')

    for col in ['PESO_CSV', 'PRECO_CSV', 'OS', 'NF', 'CODPRODUTO']:
        csv_df[col] = pd.to_numeric(csv_df[col], errors='coerce')

    margem_df = margem_df.dropna(subset=['OS', 'NF-E', 'CODPRODUTO'])
    csv_df = csv_df.dropna(subset=['OS', 'NF', 'CODPRODUTO'])

    return margem_df, csv_df


def comparar(margem_df, csv_df):

    print("\nRealizando merge...")

    df = pd.merge(
        margem_df,
        csv_df,
        left_on=['OS', 'NF-E', 'CODPRODUTO'],
        right_on=['OS', 'NF', 'CODPRODUTO'],
        how='inner'
    )

    if df.empty:
        return pd.DataFrame()

    qtde = df['QTDE AJUSTADA']
    preco = df['Preço Venda ']
    peso = df['PESO_CSV']
    preco_csv = df['PRECO_CSV']

    negativo = (qtde < 0) & (preco < 0)

    peso_ref = np.where(negativo, -np.abs(peso), np.abs(peso))
    preco_ref = np.where(negativo, -np.abs(preco_csv), np.abs(preco_csv))

    cf_ref = np.where(negativo, 'DEV', 'ESP')
    hist_ref = np.where(negativo, '68', '51')

    resultado = pd.DataFrame({
        'STATUS': np.where(
            np.isclose(qtde, peso_ref, atol=0.1) &
            np.isclose(preco, preco_ref, atol=0.01) &
            df['CF'].astype(str).str.strip().eq(cf_ref) &
            df['HISTORICO_CSV'].astype(str).str.strip().eq(hist_ref),
            'CORRETO',
            'ERRO'
        ),
        'OS': df['OS'],
        'NF': df['NF-E'],
        'COD': df['CODPRODUTO'],
        'CF': df['CF'],
        'HISTORICO': df['HISTORICO_CSV'],
        'QTDE': qtde,
        'PESO': peso,
        'PRECO': preco,
        'PRECO_CSV': preco_csv
    })

    return resultado


def salvar(df):

    if df.empty:
        print("Nenhum resultado.")
        return None

    corretos = df[df['STATUS'] == 'CORRETO']
    erros = df[df['STATUS'] == 'ERRO']

    output = r"C:\Users\DELL\Downloads\MAR x MOV.xlsx"

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        corretos.to_excel(writer, sheet_name='CORRETOS', index=False)
        erros.to_excel(writer, sheet_name='ERROS', index=False)
        df.to_excel(writer, sheet_name='TODOS', index=False)

    print("\n=== RESULTADO ===")
    print(f"Total: {len(df)}")
    print(f"Corretos: {len(corretos)}")
    print(f"Erros: {len(erros)}")

    return output


def main():
    try:
        print("Iniciando...")

        margem, csv = carregar_planilhas()
        margem, csv = limpar_dados(margem, csv)
        resultado = comparar(margem, csv)

        if resultado.empty:
            print("Sem dados para comparar.")
            return

        arquivo = salvar(resultado)

        if arquivo:
            print(f"\nArquivo gerado: {arquivo}")

    except Exception as e:
        print(f"Erro: {e}")


if __name__ == "__main__":
    main()