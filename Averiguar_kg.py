import pandas as pd
import numpy as np

def carregar_planilhas():
    """Carrega e prepara as planilhas para análise"""
    
    print("Carregando arquivo Excel...")
    margem_df = pd.read_excel(
        r"C:\Users\win11\Downloads\Margem_250929 - wapp.xlsx",
        sheet_name="Base (3,5%)",
        header=8,
        skiprows=0
    )
    
    print("Carregando arquivo CSV...")
    try:
        # Detectar separador
        with open(r"C:\Users\win11\Downloads\20250901.csv", 'r', encoding='utf-8') as f:
            first_line = f.readline()
            sep = ';' if first_line.count(';') > first_line.count(',') else ','
    
        csv_df = pd.read_csv(
            r"C:\Users\win11\Downloads\20250901.csv",
            encoding='utf-8',
            sep=sep,
            engine='python',
            on_bad_lines='skip',
            decimal=',',
            thousands='.',
            dtype={'HISTORICO': str}  # Forçar como string
        )
            
    except Exception as e:
        print(f"Erro ao carregar CSV: {e}")
        return None, None
    
    return margem_df, csv_df

def limpar_e_preparar_dados(margem_df, csv_df):
    """Limpa e prepara os dados para comparação"""
    
    # Verificar colunas disponíveis
    print("\nColunas no CSV:", list(csv_df.columns))
    print("Colunas na Margem:", list(margem_df.columns))
    
    # CORREÇÃO: Verificar se a coluna HISTORICO existe e seus valores
    if 'HISTORICO' in csv_df.columns:
        print("\nValores únicos em HISTORICO:", csv_df['HISTORICO'].unique())
        print("Total de HISTORICO vazios:", csv_df['HISTORICO'].isna().sum())
        print("Total de HISTORICO com valor:", csv_df['HISTORICO'].notna().sum())
    
    # Renomear colunas para facilitar
    csv_df = csv_df.rename(columns={
        'ROMANEIO': 'OS',
        'NOTA FISCAL': 'NF', 
        'PRODUTO': 'CODPRODUTO',
        'PESO': 'PESO_CSV',
        'UNITARIO': 'PRECO_CSV',
        'HISTORICO': 'HISTORICO_CSV'
    })
    
    # Selecionar apenas colunas necessárias da Margem
    colunas_margem = ['OS', 'NF-E', 'CODPRODUTO', 'QTDE AJUSTADA', 'Preço Venda ', 'CF']
    margem_df = margem_df[colunas_margem].copy()
    
    # Converter colunas numéricas
    margem_df['QTDE AJUSTADA'] = pd.to_numeric(margem_df['QTDE AJUSTADA'], errors='coerce')
    margem_df['Preço Venda '] = pd.to_numeric(margem_df['Preço Venda '], errors='coerce')
    
    csv_df['PESO_CSV'] = pd.to_numeric(csv_df['PESO_CSV'], errors='coerce')
    csv_df['PRECO_CSV'] = pd.to_numeric(csv_df['PRECO_CSV'], errors='coerce')
    csv_df['OS'] = pd.to_numeric(csv_df['OS'], errors='coerce')
    csv_df['NF'] = pd.to_numeric(csv_df['NF'], errors='coerce')
    csv_df['CODPRODUTO'] = pd.to_numeric(csv_df['CODPRODUTO'], errors='coerce')
    
    # Remover linhas com valores vazios nas chaves
    margem_df = margem_df.dropna(subset=['OS', 'NF-E', 'CODPRODUTO'])
    csv_df = csv_df.dropna(subset=['OS', 'NF', 'CODPRODUTO'])
    
    return margem_df, csv_df

def realizar_comparacao(margem_df, csv_df):
    """Realiza a comparação simplificada"""
    
    print("\nRealizando merge...")
    
    # Fazer merge
    merged_df = pd.merge(
        margem_df,
        csv_df,
        left_on=['OS', 'NF-E', 'CODPRODUTO'],
        right_on=['OS', 'NF', 'CODPRODUTO'],
        how='inner'
    )
    
    print(f"Registros após merge: {len(merged_df)}")
    
    if len(merged_df) == 0:
        return pd.DataFrame()
    
    # Aplicar lógica de comparação
    resultados = []
    
    for _, row in merged_df.iterrows():
        try:
            qtde = row['QTDE AJUSTADA']
            preco = row['Preço Venda ']
            peso_csv = row['PESO_CSV']
            preco_csv = row['PRECO_CSV']
            cf = row.get('CF', '')
            historico = row.get('HISTORICO_CSV', '')
            
            # Pular se valores forem NaN
            if pd.isna(qtde) or pd.isna(preco) or pd.isna(peso_csv) or pd.isna(preco_csv):
                continue
            
            # Aplicar lógica de negativos
            if qtde < 0 and preco < 0:
                peso_comparar = -abs(peso_csv)
                preco_comparar = -abs(preco_csv)
                cf_esperado = 'DEV'
                historico_esperado = '68'
            else:
                peso_comparar = abs(peso_csv)
                preco_comparar = abs(preco_csv)
                cf_esperado = 'ESP'
                historico_esperado = '51'
            
            # Verificar matches
            peso_match = abs(qtde - peso_comparar) < 0.1
            preco_match = abs(preco - preco_comparar) < 0.01
            cf_match = str(cf).strip() == cf_esperado
            historico_match = str(historico).strip() == historico_esperado
            
            # Determinar status
            if peso_match and preco_match and cf_match and historico_match:
                status = 'CORRETO'
            else:
                status = 'ERRO'
                
            resultados.append({
                'STATUS': status,
                'OS': row['OS'],
                'NF': row['NF-E'],
                'COD': row['CODPRODUTO'],
                'CF': cf,
                'HISTORICO': historico,
                'QTDE_AJUSTADA': qtde,
                'PESO': peso_csv,
                'Preço Venda': preco,
                'PRECO': preco_csv
            })
            
        except Exception as e:
            continue
    
    return pd.DataFrame(resultados)

def criar_planilha_resultados(df):
    
    if df.empty:
        print("Nenhum resultado!")
        return None
    
    # Separar corretos e erros
    corretos = df[df['STATUS'] == 'CORRETO']
    erros = df[df['STATUS'] == 'ERRO']
    
    output_path = r"C:\Users\win11\Downloads\MAR x MOV.xlsx"
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        corretos.to_excel(writer, sheet_name='CORRETOS', index=False)
        erros.to_excel(writer, sheet_name='ERROS', index=False)
        df.to_excel(writer, sheet_name='TODOS', index=False)
    
    # Estatísticas simples
    total = len(df)
    total_corretos = len(corretos)
    
    print(f"\n=== RESULTADOS ===")
    print(f"Total analisado: {total}")
    print(f"Registros corretos: {total_corretos} ({total_corretos/total*100:.1f}%)")
    print(f"Registros com erro: {total - total_corretos} ({(total-total_corretos)/total*100:.1f}%)")
    
    return output_path

def main():
    """Função principal simplificada"""
    try:
        print("Iniciando análise...")
        
        # Carregar dados
        margem_df, csv_df = carregar_planilhas()
        
        if margem_df is None or csv_df is None:
            print("Erro ao carregar arquivos!")
            return
        
        print(f"Margem: {len(margem_df)} registros")
        print(f"CSV: {len(csv_df)} registros")
        
        # Preparar dados
        margem_clean, csv_clean = limpar_e_preparar_dados(margem_df, csv_df)
        
        # Comparar
        resultados = realizar_comparacao(margem_clean, csv_clean)
        
        if resultados.empty:
            print("Nenhum registro para comparar!")
            return
        
        # Salvar resultados
        arquivo = criar_planilha_resultados(resultados)
        
        if arquivo:
            print(f"\nArquivo salvo: {arquivo}")
            
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    main()