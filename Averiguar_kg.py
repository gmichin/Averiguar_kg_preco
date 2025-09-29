import pandas as pd
import numpy as np
from pathlib import Path

def carregar_planilhas():
    """Carrega e prepara as planilhas para análise"""
    
    # Carregar arquivo Excel (Margem)
    print("Carregando arquivo Excel...")
    margem_df = pd.read_excel(
        r"C:\Users\win11\Downloads\Margem_250925 - wapp.xlsx",
        sheet_name="Base (3,5%)",
        header=8,  # Começa na linha 9 (índice 8)
        skiprows=0
    )
    
    # Carregar arquivo CSV com tratamento de erros
    print("Carregando arquivo CSV...")
    try:
        # Primeiro, vamos verificar a estrutura do CSV
        with open(r"C:\Users\win11\Downloads\20250901.csv", 'r', encoding='utf-8') as f:
            lines = f.readlines()
            print(f"Total de linhas no CSV: {len(lines)}")
            
            # Determinar separador
            first_line = lines[0]
            if first_line.count(';') > first_line.count(','):
                sep = ';'
            else:
                sep = ','
                
            print(f"Separador detectado: '{sep}'")
    
        # Carregar CSV
        csv_df = pd.read_csv(
            r"C:\Users\win11\Downloads\20250901.csv",
            encoding='utf-8',
            sep=sep,
            engine='python',
            on_bad_lines='skip',
            decimal=',',  # Adicionar esta linha para tratar decimal brasileiro
            thousands='.'  # Adicionar esta linha para tratar milhar brasileiro
        )
            
    except Exception as e:
        print(f"Erro ao carregar CSV: {e}")
        csv_df = carregar_csv_manual()
    
    return margem_df, csv_df

def carregar_csv_manual():
    """Carrega o CSV manualmente para lidar com problemas de formatação"""
    file_path = r"C:\Users\win11\Downloads\20250901.csv"
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Determinar separador
    first_line = lines[0]
    if first_line.count(';') > first_line.count(','):
        sep = ';'
    else:
        sep = ','
    
    # Processar cabeçalho
    headers = first_line.strip().split(sep)
    print(f"Cabeçalhos encontrados no CSV: {headers}")
    
    # Coletar dados
    data = []
    
    for i, line in enumerate(lines[1:], 1):
        line = line.strip()
        if not line:
            continue
            
        parts = line.split(sep)
        if len(parts) == len(headers):
            data.append(parts)
        elif len(parts) > len(headers):
            # Juntar os campos extras no último campo
            corrected = parts[:len(headers)-1] + [sep.join(parts[len(headers)-1:])]
            if len(corrected) == len(headers):
                data.append(corrected)
    
    csv_df = pd.DataFrame(data, columns=headers)
    print(f"CSV carregado manualmente: {len(csv_df)} registros")
    
    return csv_df

def limpar_e_preparar_dados(margem_df, csv_df):
    """Limpa e prepara os dados para comparação"""
    
    # Fazer uma cópia dos DataFrames
    margem_df = margem_df.copy()
    csv_df = csv_df.copy()
    
    # Verificar e imprimir colunas disponíveis
    print("\n=== COLUNAS NA PLANILHA MARGEM ===")
    for i, col in enumerate(margem_df.columns):
        print(f"{i:2d}. '{col}'")
    
    print("\n=== COLUNAS NO CSV ===")
    for i, col in enumerate(csv_df.columns):
        print(f"{i:2d}. '{col}'")
    
    # CORREÇÃO: A coluna na planilha Margem tem um espaço no final
    colunas_necessarias_margem = ['OS', 'NF-E', 'CODPRODUTO', 'QTDE AJUSTADA', 'Preço Venda ', 'CF']
    colunas_necessarias_csv = ['ROMANEIO', 'NOTA FISCAL', 'PRODUTO', 'PESO', 'UNITARIO', 'HISTORICO']
    
    print("\n=== VERIFICAÇÃO DE COLUNAS ===")
    for col in colunas_necessarias_margem:
        if col in margem_df.columns:
            print(f"✓ Coluna '{col}' encontrada na planilha Margem")
        else:
            print(f"✗ Coluna '{col}' NÃO encontrada na planilha Margem")
            colunas_similares = [c for c in margem_df.columns if col.strip().lower() in c.lower()]
            if colunas_similares:
                print(f"  Colunas similares: {colunas_similares}")
    
    for col in colunas_necessarias_csv:
        if col in csv_df.columns:
            print(f"✓ Coluna '{col}' encontrada no CSV")
        else:
            print(f"✗ Coluna '{col}' NÃO encontrada no CSV")
            colunas_similares = [c for c in csv_df.columns if col.lower() in c.lower()]
            if colunas_similares:
                print(f"  Colunas similares: {colunas_similares}")
    
    print("\n=== CONVERTENDO COLUNAS NUMÉRICAS ===")
    
    # Mapeamento de colunas CORRIGIDO
    mapeamento_colunas = {
        'OS': ['OS'],
        'NF-E': ['NF-E', 'NF_E', 'NFE'],
        'CODPRODUTO': ['CODPRODUTO', 'COD PRODUTO', 'CODPROD'],
        'QTDE AJUSTADA': ['QTDE AJUSTADA', 'QTDE_AJUSTADA', 'QTD AJUSTADA'],
        'Preço Venda': ['Preço Venda ', 'Preço Venda', 'Preco Venda', 'PRECO VENDA'],
        'CF': ['CF']
    }
    
    # Encontrar os nomes reais das colunas
    colunas_reais_margem = {}
    for col_padrao, alternativas in mapeamento_colunas.items():
        for alt in alternativas:
            if alt in margem_df.columns:
                colunas_reais_margem[col_padrao] = alt
                print(f"Usando '{alt}' para '{col_padrao}'")
                break
        if col_padrao not in colunas_reais_margem:
            print(f"AVISO: Nenhuma coluna encontrada para '{col_padrao}'")
    
    # Converter colunas da planilha Margem
    for col_padrao, col_real in colunas_reais_margem.items():
        if col_real in margem_df.columns:
            # Para Preço Venda, pode ter formatação brasileira
            if 'Preço' in col_padrao or 'Preco' in col_padrao:
                # Primeiro tentar converter diretamente
                margem_df[col_padrao] = pd.to_numeric(margem_df[col_real], errors='coerce')
                # Se muitos NaN, tentar tratar formatação brasileira
                if margem_df[col_padrao].isna().sum() > len(margem_df) * 0.5:
                    print(f"Tentando converter {col_real} com formatação brasileira...")
                    margem_df[col_padrao] = margem_df[col_real].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                    margem_df[col_padrao] = pd.to_numeric(margem_df[col_padrao], errors='coerce')
            elif col_padrao in ['QTDE AJUSTADA']:
                margem_df[col_padrao] = pd.to_numeric(margem_df[col_real], errors='coerce')
            elif col_padrao in ['CF']:
                # Manter como string para CF
                margem_df[col_padrao] = margem_df[col_real].astype(str)
            else:
                margem_df[col_padrao] = pd.to_numeric(margem_df[col_real], errors='coerce')
            
            if col_padrao not in ['CF']:
                print(f"Convertida '{col_real}' -> '{col_padrao}': {margem_df[col_padrao].notna().sum()} valores válidos")
    
    # Converter colunas do CSV - CORREÇÃO PARA FORMATAÇÃO BRASILEIRA
    print("\n=== CONVERTENDO COLUNAS DO CSV ===")
    
    # Primeiro, vamos inspecionar alguns valores das colunas problemáticas
    print("Amostra de valores da coluna PESO no CSV:")
    print(csv_df['PESO'].head(10))
    print("Amostra de valores da coluna UNITARIO no CSV:")
    print(csv_df['UNITARIO'].head(10))
    
    # Função para converter formatação brasileira
    def converter_brasileiro(valor):
        if pd.isna(valor):
            return valor
        try:
            # Se já for numérico, retorna como está
            if isinstance(valor, (int, float)):
                return float(valor)
            # Se for string, trata formatação brasileira
            if isinstance(valor, str):
                # Remove espaços e R$
                valor = valor.strip().replace('R$', '').replace(' ', '')
                # Substitui vírgula por ponto e remove pontos de milhar
                if ',' in valor and '.' in valor:
                    # Formato: 1.234,56 -> remove ponto, substitui vírgula
                    valor = valor.replace('.', '').replace(',', '.')
                elif ',' in valor:
                    # Formato: 1234,56 -> substitui vírgula
                    valor = valor.replace(',', '.')
                return float(valor)
            return float(valor)
        except:
            return np.nan
    
    # Converter colunas do CSV
    colunas_csv_para_converter = ['PESO', 'UNITARIO']
    for col in colunas_csv_para_converter:
        if col in csv_df.columns:
            print(f"Convertendo coluna '{col}'...")
            
            # Tentar conversão direta primeiro
            csv_df[col] = pd.to_numeric(csv_df[col], errors='coerce')
            validos_direto = csv_df[col].notna().sum()
            
            # Se muitos valores inválidos, tentar formatação brasileira
            if validos_direto < len(csv_df) * 0.5:
                print(f"Poucos valores válidos com conversão direta ({validos_direto}), tentando formatação brasileira...")
                csv_df[col] = csv_df[col].apply(converter_brasileiro)
            
            print(f"Coluna '{col}': {csv_df[col].notna().sum()} valores válidos")
            print(f"Valores únicos sample: {csv_df[col].dropna().head(5).tolist()}")
    
    # Também converter as colunas de chave
    for col in ['ROMANEIO', 'NOTA FISCAL', 'PRODUTO']:
        if col in csv_df.columns:
            csv_df[col] = pd.to_numeric(csv_df[col], errors='coerce')
            print(f"Coluna chave '{col}': {csv_df[col].notna().sum()} valores válidos")
    
    # Para HISTORICO no CSV, manter como string
    for col in ['HISTORICO']:
        if col in csv_df.columns:
            csv_df[col] = csv_df[col].astype(str)
            print(f"Coluna '{col}': {csv_df[col].notna().sum()} valores válidos")
    
    # Remover linhas com valores NaN nas colunas chave
    chaves_margem = [colunas_reais_margem.get(col, col) for col in ['OS', 'NF-E', 'CODPRODUTO'] if col in colunas_reais_margem]
    chaves_csv = ['ROMANEIO', 'NOTA FISCAL', 'PRODUTO']
    
    margem_original = len(margem_df)
    csv_original = len(csv_df)
    
    margem_df_clean = margem_df.dropna(subset=chaves_margem).copy()
    csv_df_clean = csv_df.dropna(subset=chaves_csv).copy()
    
    print(f"\n=== LIMPEZA DE DADOS ===")
    print(f"Margem: {margem_original} -> {len(margem_df_clean)} (removidos {margem_original - len(margem_df_clean)})")
    print(f"CSV: {csv_original} -> {len(csv_df_clean)} (removidos {csv_original - len(csv_df_clean)})")
    
    # Verificar se temos dados nas colunas de comparação
    print(f"\n=== DADOS NAS COLUNAS DE COMPARAÇÃO ===")
    print(f"QTDE AJUSTADA (Margem): {margem_df_clean['QTDE AJUSTADA'].notna().sum()} válidos")
    print(f"PESO (CSV): {csv_df_clean['PESO'].notna().sum()} válidos")
    print(f"Preço Venda (Margem): {margem_df_clean['Preço Venda'].notna().sum()} válidos")
    print(f"UNITARIO (CSV): {csv_df_clean['UNITARIO'].notna().sum()} válidos")
    print(f"CF (Margem): {margem_df_clean['CF'].notna().sum()} válidos")
    print(f"HISTORICO (CSV): {csv_df_clean['HISTORICO'].notna().sum()} válidos")
    
    return margem_df_clean, csv_df_clean, colunas_reais_margem

def determinar_historico_correto(qtde_ajustada, preco_venda):
    """Determina qual HISTORICO deveria estar baseado nos valores"""
    if qtde_ajustada < 0 and preco_venda < 0:
        return '68'
    else:
        return '51'

def verificar_cf_e_historico(qtde_ajustada, preco_venda, cf_margem, historico_csv):
    """Verifica se CF e HISTORICO estão corretos conforme a lógica de valores negativos"""
    
    # Verificar se ambos são negativos
    ambos_negativos = qtde_ajustada < 0 and preco_venda < 0
    
    # Verificar CF
    cf_correto = True
    if ambos_negativos:
        if cf_margem != 'DEV':
            cf_correto = False
    else:
        if cf_margem != 'ESP':
            cf_correto = False
    
    # Verificar HISTORICO
    historico_correto = True
    historico_esperado = determinar_historico_correto(qtde_ajustada, preco_venda)
    if historico_csv != historico_esperado:
        historico_correto = False
    
    return cf_correto, historico_correto, historico_esperado

def realizar_comparacao(margem_df, csv_df, colunas_reais_margem):
    """Realiza a comparação entre as planilhas incluindo a lógica de CF e HISTORICO"""
    
    print("\n=== REALIZANDO COMPARAÇÃO ===")
    
    # Verificar se temos as colunas necessárias para o merge
    colunas_merge_margem = ['OS', 'NF-E', 'CODPRODUTO']
    colunas_merge_csv = ['ROMANEIO', 'NOTA FISCAL', 'PRODUTO']
    
    # Verificar se todas as colunas necessárias existem
    for col in colunas_merge_margem:
        if col not in margem_df.columns:
            print(f"ERRO: Coluna '{col}' não encontrada para merge")
            return pd.DataFrame()
    
    for col in colunas_merge_csv:
        if col not in csv_df.columns:
            print(f"ERRO: Coluna '{col}' não encontrada no CSV para merge")
            return pd.DataFrame()
    
    # Mesclar os dataframes
    print("Realizando merge das planilhas...")
    merged_df = pd.merge(
        margem_df,
        csv_df,
        left_on=colunas_merge_margem,
        right_on=colunas_merge_csv,
        how='inner',
        suffixes=('_margem', '_csv')
    )
    
    print(f"Total de registros correspondentes após merge: {len(merged_df)}")
    
    if len(merged_df) == 0:
        print("AVISO: Nenhum registro correspondente encontrado!")
        return pd.DataFrame()
    
    # Verificar valores antes da comparação
    print(f"\n=== AMOSTRA DE VALORES PARA COMPARAÇÃO ===")
    print("5 primeiros registros com valores:")
    sample_data = []
    for idx, row in merged_df.head(5).iterrows():
        sample_data.append({
            'OS/ROMANEIO': row['OS'],
            'QTDE_AJUSTADA': row.get('QTDE AJUSTADA', 'N/A'),
            'PESO': row.get('PESO', 'N/A'),
            'Preço_Venda': row.get('Preço Venda', 'N/A'),
            'UNITARIO': row.get('UNITARIO', 'N/A'),
            'CF_margem': row.get('CF_margem', 'N/A'),
            'HISTORICO_csv': row.get('HISTORICO_csv', 'N/A')
        })
    
    for sample in sample_data:
        print(f"  OS {sample['OS/ROMANEIO']}: QTDE={sample['QTDE_AJUSTADA']}, PESO={sample['PESO']}, Preço={sample['Preço_Venda']}, Unitário={sample['UNITARIO']}, CF={sample['CF_margem']}, HIST={sample['HISTORICO_csv']}")
    
    # Criar lista para armazenar os resultados
    resultados = []
    erros_processamento = 0
    
    for index, row in merged_df.iterrows():
        try:
            romaneio = row['ROMANEIO']
            nfe = row['NF-E']
            codproduto = row['CODPRODUTO']
            
            # Obter valores com verificação de NaN
            qtde_ajustada = row.get('QTDE AJUSTADA', 0)
            peso = row.get('PESO', 0)
            preco_venda = row.get('Preço Venda', 0)
            unitario = row.get('UNITARIO', 0)
            cf_margem = row.get('CF_margem', '')
            historico_csv = row.get('HISTORICO_csv', '')
            
            # Verificar se os valores são NaN
            if pd.isna(qtde_ajustada) or pd.isna(peso) or pd.isna(preco_venda) or pd.isna(unitario):
                erros_processamento += 1
                continue
            
            # Converter para float para comparação
            qtde_ajustada = float(qtde_ajustada)
            peso = float(peso)
            preco_venda = float(preco_venda)
            unitario = float(unitario)
            
            # APLICAR LÓGICA DE VALORES NEGATIVOS
            # Se ambos são negativos na planilha Margem, converter valores do CSV para negativos
            if qtde_ajustada < 0 and preco_venda < 0:
                peso_comparar = -abs(peso)  # Converter para negativo
                unitario_comparar = -abs(unitario)  # Converter para negativo
            else:
                peso_comparar = abs(peso)  # Manter positivo
                unitario_comparar = abs(unitario)  # Manter positivo
            
            # Verificar correspondência de peso (com tolerância maior para peso)
            peso_match = abs(qtde_ajustada - peso_comparar) < 0.1  # Aumentei a tolerância para 0.1
            
            # Verificar correspondência de preço (com tolerância para centavos)
            preco_match = abs(preco_venda - unitario_comparar) < 0.01
            
            # Verificar CF e HISTORICO
            cf_correto, historico_correto, historico_esperado = verificar_cf_e_historico(
                qtde_ajustada, preco_venda, cf_margem, historico_csv
            )
            
            # Classificar o registro
            status = 'Corretos'
            erros = []
            
            if not peso_match:
                erros.append('Peso')
            if not preco_match:
                erros.append('Preco')
            if not cf_correto:
                erros.append('CF')
            if not historico_correto:
                erros.append('HISTORICO')
            
            if erros:
                status = '_'.join(erros) + '_erro'
            
            resultados.append({
                'ROMANEIO': romaneio,
                'NF-E': nfe,
                'COD': codproduto,
                'QTDE_AJUSTADA': qtde_ajustada,
                'PESO': peso,
                'PESO_COMPARAR': peso_comparar,  # Valor usado na comparação
                'Preco_Venda': preco_venda,
                'UNITARIO': unitario,
                'UNITARIO_COMPARAR': unitario_comparar,  # Valor usado na comparação
                'CF_margem': cf_margem,
                'HISTORICO_csv': historico_csv,
                'HISTORICO_ESPERADO': historico_esperado,
                'Status': status,
                'CF_correto': cf_correto,
                'HISTORICO_correto': historico_correto
            })
            
        except Exception as e:
            erros_processamento += 1
            if erros_processamento <= 5:  # Mostrar apenas os primeiros 5 erros
                print(f"Erro ao processar linha {index}: {e}")
                print(f"  Valores: QTDE={row.get('QTDE AJUSTADA', 'N/A')}, PESO={row.get('PESO', 'N/A')}, Preço={row.get('Preço Venda', 'N/A')}, Unitário={row.get('UNITARIO', 'N/A')}")
            continue
    
    if erros_processamento > 0:
        print(f"Total de erros de processamento: {erros_processamento}")
    
    return pd.DataFrame(resultados)

def criar_planilha_resultados(df_resultados):
    """Cria a planilha Excel com as abas separadas incluindo CF e HISTORICO"""
    
    if df_resultados.empty:
        print("Nenhum dado para salvar!")
        return None
    
    # Filtrar dados por categoria
    corretos_df = df_resultados[df_resultados['Status'] == 'Corretos']
    peso_erro_df = df_resultados[df_resultados['Status'].str.contains('Peso')]
    preco_erro_df = df_resultados[df_resultados['Status'].str.contains('Preco')]
    cf_erro_df = df_resultados[df_resultados['Status'].str.contains('CF')]
    historico_erro_df = df_resultados[df_resultados['Status'].str.contains('HISTORICO')]
    multiplos_erros_df = df_resultados[df_resultados['Status'].str.contains('_') & ~df_resultados['Status'].isin(['Corretos'])]
    
    # Selecionar apenas as colunas necessárias (incluindo CF e HISTORICO)
    colunas_finais = [
        'ROMANEIO', 'NF-E', 'COD', 
        'QTDE_AJUSTADA', 'PESO', 'PESO_COMPARAR',
        'Preco_Venda', 'UNITARIO', 'UNITARIO_COMPARAR',
        'CF_margem', 'HISTORICO_csv', 'HISTORICO_ESPERADO',
        'CF_correto', 'HISTORICO_correto', 'Status'
    ]
    
    # Criar arquivo Excel
    output_path = r"C:\Users\win11\Downloads\Resultado_Averiguacao.xlsx"
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Aba Corretos
            if not corretos_df.empty:
                corretos_df[colunas_finais].to_excel(writer, sheet_name='Corretos', index=False)
                print(f"Aba 'Corretos': {len(corretos_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='Corretos', index=False)
                print("Aba 'Corretos': 0 registros")
            
            # Aba Peso erro
            if not peso_erro_df.empty:
                peso_erro_df[colunas_finais].to_excel(writer, sheet_name='Peso_erro', index=False)
                print(f"Aba 'Peso_erro': {len(peso_erro_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='Peso_erro', index=False)
                print("Aba 'Peso_erro': 0 registros")
            
            # Aba Preco erro
            if not preco_erro_df.empty:
                preco_erro_df[colunas_finais].to_excel(writer, sheet_name='Preco_erro', index=False)
                print(f"Aba 'Preco_erro': {len(preco_erro_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='Preco_erro', index=False)
                print("Aba 'Preco_erro': 0 registros")
            
            # Aba CF erro
            if not cf_erro_df.empty:
                cf_erro_df[colunas_finais].to_excel(writer, sheet_name='CF_erro', index=False)
                print(f"Aba 'CF_erro': {len(cf_erro_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='CF_erro', index=False)
                print("Aba 'CF_erro': 0 registros")
            
            # Aba HISTORICO erro
            if not historico_erro_df.empty:
                historico_erro_df[colunas_finais].to_excel(writer, sheet_name='HISTORICO_erro', index=False)
                print(f"Aba 'HISTORICO_erro': {len(historico_erro_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='HISTORICO_erro', index=False)
                print("Aba 'HISTORICO_erro': 0 registros")
            
            # Aba para múltiplos erros
            if not multiplos_erros_df.empty:
                multiplos_erros_df[colunas_finais].to_excel(writer, sheet_name='Multiplos_Erros', index=False)
                print(f"Aba 'Multiplos_Erros': {len(multiplos_erros_df)} registros")
            else:
                pd.DataFrame(columns=colunas_finais).to_excel(writer, sheet_name='Multiplos_Erros', index=False)
                print("Aba 'Multiplos_Erros': 0 registros")
            
            # Aba com todos os registros
            df_resultados[colunas_finais].to_excel(writer, sheet_name='Todos_Registros', index=False)
            print(f"Aba 'Todos_Registros': {len(df_resultados)} registros")
        
        return output_path
        
    except Exception as e:
        print(f"Erro ao salvar arquivo Excel: {e}")
        return None

def main():
    """Função principal"""
    try:
        print("Iniciando processo de averiguação...")
        
        # Carregar planilhas
        margem_df, csv_df = carregar_planilhas()
        
        print(f"Registros na planilha Margem: {len(margem_df)}")
        print(f"Registros na planilha CSV: {len(csv_df)}")
        
        if len(margem_df) == 0 or len(csv_df) == 0:
            print("ERRO: Uma das planilhas está vazia!")
            return
        
        # Limpar e preparar dados
        margem_df_clean, csv_df_clean, colunas_reais = limpar_e_preparar_dados(margem_df, csv_df)
        
        # Realizar comparação
        resultados_df = realizar_comparacao(margem_df_clean, csv_df_clean, colunas_reais)
        
        if resultados_df.empty:
            print("Nenhum resultado para processar!")
            return
        
        print(f"Total de registros comparados: {len(resultados_df)}")
        
        # Criar planilha de resultados
        output_path = criar_planilha_resultados(resultados_df)
        
        if output_path:
            # Estatísticas
            total = len(resultados_df)
            corretos = len(resultados_df[resultados_df['Status'] == 'Corretos'])
            peso_erro = len(resultados_df[resultados_df['Status'].str.contains('Peso')])
            preco_erro = len(resultados_df[resultados_df['Status'].str.contains('Preco')])
            cf_erro = len(resultados_df[resultados_df['Status'].str.contains('CF')])
            historico_erro = len(resultados_df[resultados_df['Status'].str.contains('HISTORICO')])
            
            print("\n=== RESULTADO DA AVERIGUAÇÃO ===")
            print(f"Total de registros analisados: {total}")
            if total > 0:
                print(f"Registros corretos: {corretos} ({corretos/total*100:.1f}%)")
                print(f"Erros de peso: {peso_erro} ({peso_erro/total*100:.1f}%)")
                print(f"Erros de preço: {preco_erro} ({preco_erro/total*100:.1f}%)")
                print(f"Erros de CF: {cf_erro} ({cf_erro/total*100:.1f}%)")
                print(f"Erros de HISTORICO: {historico_erro} ({historico_erro/total*100:.1f}%)")
            
            # Mostrar exemplos de registros com valores negativos
            negativos_df = resultados_df[
                (resultados_df['QTDE_AJUSTADA'] < 0) & 
                (resultados_df['Preco_Venda'] < 0)
            ]
            if len(negativos_df) > 0:
                print(f"\nRegistros com valores negativos: {len(negativos_df)}")
                print("Exemplo de registros negativos:")
                for idx, row in negativos_df.head(3).iterrows():
                    print(f"  OS {row['ROMANEIO']}: QTDE={row['QTDE_AJUSTADA']}, Preço={row['Preco_Venda']}, CF={row['CF_margem']}, HIST_esperado={row['HISTORICO_ESPERADO']}, HIST_atual={row['HISTORICO_csv']}")
            
            print(f"\nArquivo salvo em: {output_path}")
        
    except Exception as e:
        print(f"Erro durante o processo: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()