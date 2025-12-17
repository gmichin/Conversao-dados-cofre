import pandas as pd
import os
from datetime import datetime
from openpyxl.utils import get_column_letter

def processar_relatorio_nfe_final(caminho_arquivo):
    """
    Processa o relatório de NF-e baseado na estrutura identificada.
    """
    print(f"Processando: {caminho_arquivo}")
    print("-" * 60)
    
    try:
        # Ler o arquivo sem cabeçalho
        df = pd.read_excel(caminho_arquivo, header=None, dtype=str)
        print(f"Total de linhas: {len(df)}")
        print(f"Total de colunas: {df.shape[1]}")
        
        # Lista para armazenar os dados processados
        dados_formatados = []
        
        linha = 2  # Começar na linha 2 (onde estão os dados da primeira nota)
        
        while linha < len(df):
            # Verificar se é uma linha de cabeçalho de nota
            # Na linha do cabeçalho, coluna 0 tem número da nota, coluna 9 tem valor da nota
            if (pd.notna(df.iloc[linha, 0]) and 
                str(df.iloc[linha, 0]).strip().isdigit() and
                pd.notna(df.iloc[linha, 9]) and 
                ',' in str(df.iloc[linha, 9])):  # Valor tem vírgula decimal
                
                # DADOS DO CABEÇALHO DA NOTA (linha atual)
                nota_numero = str(df.iloc[linha, 0]).strip()
                
                # Tipo Oper. - usando da coluna 1
                tipo_oper = str(df.iloc[linha, 1]).strip() if pd.notna(df.iloc[linha, 1]) else ""
                # Limpar "1 - Saída" para ficar só "Saída" ou similar
                if " - " in tipo_oper:
                    tipo_oper = tipo_oper.split(" - ")[1].strip()
                
                natureza_op = str(df.iloc[linha, 2]).strip() if pd.notna(df.iloc[linha, 2]) else ""
                cnpj_dest = str(df.iloc[linha, 3]).strip() if pd.notna(df.iloc[linha, 3]) else ""
                razao_dest = str(df.iloc[linha, 4]).strip() if pd.notna(df.iloc[linha, 4]) else ""
                valor_total_str = str(df.iloc[linha, 9]).strip() if pd.notna(df.iloc[linha, 9]) else ""
                cnpj_emit = str(df.iloc[linha, 6]).strip() if pd.notna(df.iloc[linha, 6]) else ""
                razao_emit = str(df.iloc[linha, 7]).strip() if pd.notna(df.iloc[linha, 7]) else ""
                data_emissao = str(df.iloc[linha, 10]).strip() if pd.notna(df.iloc[linha, 10]) else ""
                
                # Formatar data (remover hora)
                if ' ' in data_emissao:
                    data_emissao = data_emissao.split(' ')[0]
                
                # Converter valor total para numérico (formato brasileiro)
                valor_total_numerico = None
                if valor_total_str:
                    try:
                        # Remover pontos de milhar, substituir vírgula por ponto
                        valor_limpo = valor_total_str.replace('.', '').replace(',', '.')
                        valor_total_numerico = float(valor_limpo)
                    except:
                        valor_total_numerico = None
                
                # Agora procurar a linha do PRODUTO (2 linhas abaixo)
                linha_produto = linha + 2
                
                if linha_produto < len(df):
                    # Verificar se é linha de produto
                    # A descrição do produto está na COLUNA 1 (não na 0)
                    if pd.notna(df.iloc[linha_produto, 1]):
                        desc_prod = str(df.iloc[linha_produto, 1]).strip()
                        
                        # Verificar se não é cabeçalho
                        if desc_prod.lower() not in ["desc prod", "descrição", "produto"]:
                            # Buscar CFOP - deve estar na coluna 13
                            cfop = ""
                            if linha_produto < len(df) and df.shape[1] > 13:
                                if pd.notna(df.iloc[linha_produto, 13]):
                                    cfop_raw = str(df.iloc[linha_produto, 13]).strip()
                                    # Limpar e extrair apenas números
                                    cfop = ''.join(filter(str.isdigit, cfop_raw))
                                    # Pegar apenas os primeiros 4 dígitos
                                    if len(cfop) >= 4:
                                        cfop = cfop[:4]
                            
                            # Se não encontrou CFOP na coluna 13, tentar outras
                            if not cfop:
                                for col in range(10, min(18, df.shape[1])):
                                    if pd.notna(df.iloc[linha_produto, col]):
                                        cfop_raw = str(df.iloc[linha_produto, col]).strip()
                                        if cfop_raw.isdigit() and len(cfop_raw) == 4:
                                            cfop = cfop_raw
                                            break
                            
                            # Adicionar à lista
                            dados_formatados.append({
                                'Nº da Nota': nota_numero,
                                'Descrição do Produto': desc_prod,  # Usando a descrição do produto
                                'Natureza Operação': natureza_op,
                                'CNPJ Destinatário': cnpj_dest,
                                'Razão Social Destinatário': razao_dest,
                                'Valor Total': valor_total_numerico,  # Usando valor numérico
                                'Valor Total Texto': valor_total_str,  # Mantendo também o texto original
                                'CNPJ Emitente': cnpj_emit,
                                'Razão Social Emitente': razao_emit,
                                'Emissão': data_emissao,
                                'CFOP': cfop
                            })
                            
                            print(f"✓ Nota {nota_numero}: {desc_prod[:40]}... | Valor: {valor_total_str} | CFOP: {cfop}")
                
                # Avançar 3 linhas (cabeçalho nota + linha cabeçalho produto + linha produto)
                linha += 3
            else:
                # Se não é linha de nota, avançar 1 linha
                linha += 1
        
        print(f"\nTotal de notas processadas: {len(dados_formatados)}")
        
        if not dados_formatados:
            print("❌ Nenhuma nota processada!")
            return None, None
        
        # Criar DataFrame
        df_resultado = pd.DataFrame(dados_formatados)
        
        # Converter tipos de dados
        # 1. Nº da Nota para inteiro
        df_resultado['Nº da Nota'] = pd.to_numeric(df_resultado['Nº da Nota'], errors='coerce').astype('Int64')
        
        # 2. CFOP para inteiro
        df_resultado['CFOP'] = pd.to_numeric(df_resultado['CFOP'], errors='coerce').astype('Int64')
        
        # 3. Data para datetime
        df_resultado['Emissão'] = pd.to_datetime(df_resultado['Emissão'], errors='coerce')
        
        # Calcular soma total (já está numérico)
        soma_total = df_resultado['Valor Total'].sum()
        print(f"✓ Soma total dos valores: R$ {soma_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        # Criar DataFrame final para exportação com colunas numéricas
        df_final = pd.DataFrame({
            'Nº da Nota': df_resultado['Nº da Nota'],
            'Descrição do Produto': df_resultado['Descrição do Produto'],
            'Natureza Operação': df_resultado['Natureza Operação'],
            'CNPJ Destinatário': df_resultado['CNPJ Destinatário'],
            'Razão Social Destinatário': df_resultado['Razão Social Destinatário'],
            'Valor Total': df_resultado['Valor Total'],  # Já é numérico
            'CNPJ Emitente': df_resultado['CNPJ Emitente'],
            'Razão Social Emitente': df_resultado['Razão Social Emitente'],
            'Emissão': df_resultado['Emissão'].dt.strftime('%Y-%m-%d'),  # Formata sem hora
            'CFOP': df_resultado['CFOP']
        })
        
        # Criar caminho para o novo arquivo
        caminho_pasta = os.path.dirname(caminho_arquivo)
        nome_arquivo = os.path.basename(caminho_arquivo)
        nome_sem_ext = os.path.splitext(nome_arquivo)[0]
        novo_nome = f"{nome_sem_ext}_FORMATADO_FINAL.xlsx"
        novo_caminho = os.path.join(caminho_pasta, novo_nome)
        
        # Salvar o novo arquivo Excel com formatação numérica
        with pd.ExcelWriter(novo_caminho, engine='openpyxl') as writer:
            # Adicionar título do relatório
            cabecalho_df = pd.DataFrame([['Relatório XML - 17/12/2025']])
            cabecalho_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
            
            # Adicionar "NF-E"
            nfe_df = pd.DataFrame([['NF-E']])
            nfe_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=1)
            
            # Adicionar os dados formatados
            df_final.to_excel(writer, sheet_name='Sheet1', index=False, startrow=2)
            
            # Obter a planilha para aplicar formatação
            worksheet = writer.sheets['Sheet1']
            
            # Aplicar formatação numérica para a coluna "Valor Total"
            # Encontrar a coluna "Valor Total" (coluna F, que é a 6ª coluna)
            coluna_valor_total = 6  # Coluna F (0-based index seria 5, mas no Excel é coluna F)
            
            # Formatar todas as células da coluna Valor Total como número com 2 casas decimais
            from openpyxl.styles import numbers
            
            for row in range(3, len(df_final) + 3):  # Começar na linha 3 (após cabeçalhos)
                cell = worksheet.cell(row=row, column=coluna_valor_total)
                # Aplicar formato de número brasileiro: #.##0,00
                cell.number_format = '#.##0,00'
            
            # Ajustar largura das colunas
            for col in worksheet.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"\n{'='*60}")
        print("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
        print('='*60)
        print(f"📊 Total de notas processadas: {len(df_final)}")
        print(f"💰 Soma total: R$ {soma_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        print(f"💾 Arquivo salvo em: {novo_caminho}")
        
        print(f"\n📋 Primeiras 5 notas do relatório (com valores numéricos):")
        print(df_final.head().to_string(index=False, formatters={
            'Valor Total': lambda x: f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else ''
        }))
        
        # Mostrar informações sobre os tipos de dados
        print(f"\n📝 Tipos de dados das colunas:")
        print(f"  • Nº da Nota: {df_final['Nº da Nota'].dtype}")
        print(f"  • Valor Total: {df_final['Valor Total'].dtype}")
        print(f"  • CFOP: {df_final['CFOP'].dtype}")
        print(f"  • Emissão: {type(df_final['Emissão'].iloc[0])}")
        
        return novo_caminho, df_final
        
    except Exception as e:
        print(f"❌ Erro no processamento: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def main():
    """
    Função principal
    """
    print("=" * 60)
    print("CONVERSOR DE RELATÓRIO NFE - VOG ALIMENTOS")
    print("=" * 60)
    print("✓ Valores serão exportados como números no Excel")
    print("✓ Formatação brasileira: 173.002,50")
    print("=" * 60)
    
    # Caminho do arquivo
    caminho_arquivo = r"C:\Users\win11\Downloads\RelatorioNFe-17-12-25 153350.xlsx"
    
    # Verificar se arquivo existe
    if not os.path.exists(caminho_arquivo):
        print(f"❌ Arquivo não encontrado: {caminho_arquivo}")
        print("\nPor favor, verifique:")
        print("1. O caminho está correto?")
        print("2. O arquivo está na pasta Downloads?")
        print("3. O nome do arquivo está exatamente igual?")
        return
    
    print("\n" + "=" * 60)
    print("INICIANDO PROCESSAMENTO...")
    print("=" * 60)
    
    # Processar o arquivo
    novo_caminho, df_resultado = processar_relatorio_nfe_final(caminho_arquivo)
    
    if df_resultado is not None:
        print("\n" + "=" * 60)
        print("📈 RESUMO FINAL")
        print("=" * 60)
        print(f"✅ Processamento concluído com sucesso!")
        print(f"📁 Arquivo original: {caminho_arquivo}")
        print(f"📁 Arquivo formatado: {novo_caminho}")
        print(f"📊 Total de registros: {len(df_resultado)}")
        
        # Verificar se os valores são numéricos
        if df_resultado['Valor Total'].dtype in ['float64', 'int64']:
            print("✓ Coluna 'Valor Total' está como numérica")
            print(f"✓ Soma total calculável: R$ {df_resultado['Valor Total'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            print("⚠️ Coluna 'Valor Total' NÃO está como numérica")
        
        # Mostrar exemplo
        print(f"\n📋 Exemplo dos primeiros registros:")
        sample = df_resultado.head(3).copy()
        # Formatar a exibição dos valores
        sample_display = sample.copy()
        sample_display['Valor Total'] = sample_display['Valor Total'].apply(
            lambda x: f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) else ''
        )
        print(sample_display.to_string(index=False))
        
        print(f"\n💡 Dica: No Excel, a coluna 'Valor Total' aparecerá como:")
        print(f"   • Números que podem ser somados")
        print(f"   • Formato brasileiro (#.##0,00)")
        print(f"   • Você pode usar fórmulas como =SOMA()")
    else:
        print("\n❌ Falha no processamento do arquivo.")

if __name__ == "__main__":
    main()