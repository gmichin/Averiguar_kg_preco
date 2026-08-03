import pandas as pd
import numpy as np
import chardet

def carregar_planilhas():
    """Carrega e prepara as planilhas para análise"""
    
    print("Carregando arquivo Excel - aba FEC_PQ...")
    margem_df = pd.read_excel(
        r"C:\Users\win11\Downloads\260718_MRG.xlsx",
        sheet_name="FEC_PQ",
        header=9,  # A10 corresponde à linha 10 (0-index: linha 9)
        skiprows=0
    )
    
    print("Carregando arquivo CSV...")
    try:
        # Detectar codificação
        csv_path = r"S:\hor\excel\20260701.csv"
        with open(csv_path, 'rb') as f:
            raw_data = f.read()
            encoding_result = chardet.detect(raw_data)
            encoding = encoding_result['encoding']
            print(f"Codificação detectada: {encoding} (confiança: {encoding_result['confidence']:.2f})")
        
        # Se a confiança for baixa, tentar encodings comuns
        # Sempre testar encodings comuns também
        encodings_to_try = [
            encoding,
            'cp1252',
            'latin-1',
            'iso-8859-1',
            'utf-8'
        ]

        # remover duplicados
        encodings_to_try = list(dict.fromkeys(encodings_to_try))

        
        # Detectar separador
        for enc in encodings_to_try:
            try:
                with open(csv_path, 'r', encoding=enc) as f:
                    first_line = f.readline()
                    sep = ';' if first_line.count(';') > first_line.count(',') else ','
                    print(f"Separador detectado: '{sep}' usando encoding: {enc}")
                
                csv_df = pd.read_csv(
                    csv_path,
                    encoding=enc,
                    sep=sep,
                    engine='python',
                    on_bad_lines='skip',
                    decimal=',',
                    thousands='.',
                    dtype={'HISTORICO': str}  # Forçar como string
                )
                print(f"CSV carregado com sucesso usando encoding: {enc}")
                break
                
            except UnicodeDecodeError:
                print(f"Falha com encoding {enc}, tentando próximo...")
                continue
            except Exception as e:
                print(f"Erro com encoding {enc}: {e}")
                continue
        else:
            # Se nenhum encoding funcionou, tentar sem especificar encoding
            print("Tentando carregar sem especificar encoding...")
            csv_df = pd.read_csv(
                csv_path,
                sep=sep,
                engine='python',
                on_bad_lines='skip',
                decimal=',',
                thousands='.',
                dtype={'HISTORICO': str}
            )
            
    except Exception as e:
        print(f"Erro ao carregar CSV: {e}")
        return None, None
    
    return margem_df, csv_df

def limpar_e_preparar_dados(margem_df, csv_df):
    """Limpa e prepara os dados para comparação"""
    
    # Verificar colunas disponíveis
    print("\nColunas no CSV:", list(csv_df.columns))
    print("Colunas na Margem (FEC_PQ):", list(margem_df.columns))
    
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
    
    # Renomear colunas da margem para padronizar (FEC_PQ)
    # Mapeamento das colunas da FEC_PQ para os nomes esperados
    colunas_margem_map = {
        'ROMANEIO': 'OS',           # ROMANEIO na FEC_PQ é equivalente a OS
        'NF-E': 'NF-E',             # Mantém NF-E
        'CODPRODUTO': 'CODPRODUTO',  # Mantém CODPRODUTO
        'QTDE AJUSTADA': 'QTDE AJUSTADA',
        'PRECO VENDA': 'Preço Venda ',  # PRECO VENDA na FEC_PQ
        'CF': 'CF'                   # CF na FEC_PQ
    }
    
    # Verificar quais colunas existem no DataFrame
    colunas_existentes = [col for col in colunas_margem_map.keys() if col in margem_df.columns]
    colunas_faltantes = [col for col in colunas_margem_map.keys() if col not in margem_df.columns]
    
    if colunas_faltantes:
        print(f"\nATENÇÃO: Colunas não encontradas na FEC_PQ: {colunas_faltantes}")
        print("Colunas disponíveis na FEC_PQ:", list(margem_df.columns))
    
    # Selecionar apenas colunas necessárias da Margem usando os nomes originais da FEC_PQ
    if colunas_existentes:
        margem_df = margem_df[colunas_existentes].copy()
        
        # Renomear as colunas para os nomes padronizados
        margem_df = margem_df.rename(columns=colunas_margem_map)
    else:
        print("ERRO: Nenhuma coluna esperada foi encontrada na aba FEC_PQ!")
        return None, None
    
    # Converter colunas numéricas
    if 'QTDE AJUSTADA' in margem_df.columns:
        margem_df['QTDE AJUSTADA'] = pd.to_numeric(margem_df['QTDE AJUSTADA'], errors='coerce')
    else:
        print("ERRO: Coluna 'QTDE AJUSTADA' não encontrada!")
        return None, None
    
    if 'Preço Venda ' in margem_df.columns:
        margem_df['Preço Venda '] = pd.to_numeric(margem_df['Preço Venda '], errors='coerce')
    else:
        print("ERRO: Coluna 'Preço Venda ' não encontrada!")
        return None, None
    
    if 'CF' in margem_df.columns:
        margem_df['CF'] = margem_df['CF'].astype(str).str.strip()
    
    csv_df['PESO_CSV'] = pd.to_numeric(csv_df['PESO_CSV'], errors='coerce')
    csv_df['PRECO_CSV'] = pd.to_numeric(csv_df['PRECO_CSV'], errors='coerce')
    csv_df['OS'] = pd.to_numeric(csv_df['OS'], errors='coerce')
    csv_df['NF'] = pd.to_numeric(csv_df['NF'], errors='coerce')
    csv_df['CODPRODUTO'] = pd.to_numeric(csv_df['CODPRODUTO'], errors='coerce')
    
    # Remover linhas com valores vazios nas chaves
    if 'OS' in margem_df.columns and 'NF-E' in margem_df.columns and 'CODPRODUTO' in margem_df.columns:
        margem_df = margem_df.dropna(subset=['OS', 'NF-E', 'CODPRODUTO'])
    else:
        print("ERRO: Colunas necessárias para merge não encontradas na margem!")
        return None, None
    
    csv_df = csv_df.dropna(subset=['OS', 'NF', 'CODPRODUTO'])
    
    print(f"\nApós limpeza - Margem (FEC_PQ): {len(margem_df)} registros")
    print(f"Após limpeza - CSV: {len(csv_df)} registros")
    
    # Mostrar exemplo dos dados
    if len(margem_df) > 0:
        print("\nExemplo de dados da Margem (FEC_PQ):")
        print(margem_df.head(2))
    
    return margem_df, csv_df

def realizar_comparacao(margem_df, csv_df):
    """Realiza a comparação simplificada"""
    
    print("\nRealizando merge...")
    
    # Verificar se as colunas necessárias existem
    merge_cols_margem = ['OS', 'NF-E', 'CODPRODUTO']
    merge_cols_csv = ['OS', 'NF', 'CODPRODUTO']
    
    for col in merge_cols_margem:
        if col not in margem_df.columns:
            print(f"ERRO: Coluna '{col}' não encontrada na margem para merge!")
            return pd.DataFrame()
    
    # Fazer merge
    merged_df = pd.merge(
        margem_df,
        csv_df,
        left_on=merge_cols_margem,
        right_on=merge_cols_csv,
        how='inner'
    )
    
    print(f"Registros após merge: {len(merged_df)}")
    
    if len(merged_df) == 0:
        print("Nenhum registro correspondente encontrado!")
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
    """Cria planilha com resultados"""
    
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
        print("Iniciando análise com aba FEC_PQ...")
        
        # Carregar dados
        margem_df, csv_df = carregar_planilhas()
        
        if margem_df is None or csv_df is None:
            print("Erro ao carregar arquivos!")
            return
        
        print(f"Margem (FEC_PQ): {len(margem_df)} registros")
        print(f"CSV: {len(csv_df)} registros")
        
        # Preparar dados
        margem_clean, csv_clean = limpar_e_preparar_dados(margem_df, csv_df)
        
        if margem_clean is None or csv_clean is None:
            print("Erro ao preparar os dados!")
            return
        
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
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()