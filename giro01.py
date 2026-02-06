import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

# Configuração da página Streamlit
st.set_page_config(page_title="Análise de Giro e Estoque", layout="wide")

def analisar_desempenho(df):
    """
    Executa cálculos de Giro, Cobertura e Matriz de Desempenho.
    """
    # 1. Giro de Estoque (Frequência de renovação)
    df['Giro'] = df['CMV Fat.'] / df['Estoque Custo Real'].replace(0, np.nan)
    
    # 2. Cobertura de Estoque (Dias de vida)
    df['Cobertura_Dias'] = df['Qtd. Estoque'] / df['Qtd. Média Líq.'].replace(0, np.nan)
    
    # 3. Classificação ABC (Baseada no Faturamento)
    df = df.sort_values(by='Faturamento Líquido', ascending=False)
    df['Fat_Acumulado'] = df['Faturamento Líquido'].cumsum() / df['Faturamento Líquido'].sum()
    
    def classify_abc(row):
        if row['Fat_Acumulado'] <= 0.8: 
            return 'A'
        elif row['Fat_Acumulado'] <= 0.95: 
            return 'B'
        else: 
            return 'C'
    
    df['Classe_ABC'] = df.apply(classify_abc, axis=1)
    
    # 4. Matriz Giro vs Margem
    mediana_giro = df['Giro'].median()
    mediana_margem = df['%Margem'].median()
    
    def classify_matrix(row):
        if row['Giro'] >= mediana_giro and row['%Margem'] >= mediana_margem:
            return 'Estrela (Alto Giro/Alta Margem)'
        elif row['Giro'] >= mediana_giro and row['%Margem'] < mediana_margem:
            return 'Boi Leiteiro (Alto Giro/Baixa Margem)'
        elif row['Giro'] < mediana_giro and row['%Margem'] >= mediana_margem:
            return 'Problema (Baixo Giro/Alta Margem)'
        else:
            return 'Mico (Baixo Giro/Baixa Margem)'
            
    df['Status_Estrategico'] = df.apply(classify_matrix, axis=1)
    
    return df

def main():
    # Estilo CSS para reduzir fontes e melhorar o aproveitamento de espaço
    st.markdown("""
        <style>
            /* Reduz o tamanho das métricas (KPIs) */
            [data-testid="stMetricValue"] {
                font-size: 0.85rem !important;
            }
            [data-testid="stMetricLabel"] {
                font-size: 0.85rem !important;
            }
            /* Ajusta o tamanho da fonte global e tabelas */
            .stDataFrame, div[data-testid="stTable"] {
                font-size: 12px !important;
            }
            /* Reduz margens do cabeçalho */
            .main .block-container {
                padding-top: 2rem;
            }
            h1 {
                font-size: 2rem !important;
            }
            h2 {
                font-size: 1.5rem !important;
            }
            h3 {
                font-size: 1.2rem !important;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title("📊 Painel de Giro e Desempenho de Produtos")
    st.markdown("Carregue o seu ficheiro Excel para processar os indicadores de stock automaticamente.")

    # Sidebar para Upload
    st.sidebar.header("Configurações")
    uploaded_file = st.sidebar.file_uploader("Escolha um ficheiro Excel (.xlsx)", type="xlsx")

    if uploaded_file is not None:
        try:
            # Lendo o arquivo carregado
            df_raw = pd.read_excel(uploaded_file)
            
            # Validação de Colunas
            colunas_necessarias = [
                'Referência', 'Faturamento Líquido', 'Qtd. Média Líq.', 
                '%Margem', 'CMV Fat.', 'Qtd. Estoque', 'Estoque Custo Real'
            ]
            
            colunas_faltantes = [col for col in colunas_necessarias if col not in df_raw.columns]
            
            if colunas_faltantes:
                st.error(f"Erro: As seguintes colunas não foram encontradas: {colunas_faltantes}")
                return

            # Processamento
            with st.spinner('A analisar dados...'):
                df_final = analisar_desempenho(df_raw)

            # --- KPIs de Resumo ---
            st.subheader("Indicadores de Desempenho")
            
            # Linha 1: Visão Financeira, Giro e Volume (Pares)
            m1, m2, m3, m4, m5 = st.columns(5)
            
            fat_total = df_final['Faturamento Líquido'].sum()
            estoque_total_valor = df_final['Estoque Custo Real'].sum()
            estoque_total_pares = df_final['Qtd. Estoque'].sum()
            giro_medio = df_final['Giro'].mean()
            margem_media = df_final['%Margem'].mean()
            
            m1.metric("Faturamento Total", f"R$ {fat_total:,.2f}")
            m2.metric("Stock Total (Custo)", f"R$ {estoque_total_valor:,.2f}")
            m3.metric("Stock Total (Pares)", f"{estoque_total_pares:,.0f}")
            m4.metric("Giro Médio (Mês)", f"{giro_medio:.2f}")
            m5.metric("Margem Média", f"{margem_media:.1f}%")

            # Linha 2: Status Operacional
            st.write("---")
            col1, col2, col3 = st.columns(3)
            
            rupturas = df_final[(df_final['Classe_ABC'] == 'A') & (df_final['Cobertura_Dias'] < 7)]
            micos = df_final[df_final['Status_Estrategico'].str.contains('Mico')]
            
            col1.metric("Itens Classe A", len(df_final[df_final['Classe_ABC'] == 'A']))
            col2.metric("Risco de Ruptura (A)", len(rupturas), delta_color="inverse")
            col3.metric("Itens 'Mico'", len(micos), delta_color="inverse")

            # --- Tabelas e Filtros ---
            st.divider()
            st.subheader("Visualização dos Dados Analisados")
            
            # Opção de filtro por Classe ABC
            filtro_abc = st.multiselect("Filtrar por Classe ABC", options=['A', 'B', 'C'], default=['A', 'B', 'C'])
            df_display = df_final[df_final['Classe_ABC'].isin(filtro_abc)].copy()
            
            # Renomeando colunas para a visualização
            df_display = df_display.rename(columns={'Qtd. Estoque': 'Stock (Pares)'})
            
            st.dataframe(df_display[['Referência', 'Classe_ABC', 'Giro', 'Cobertura_Dias', 'Status_Estrategico', 'Faturamento Líquido', '%Margem', 'Stock (Pares)', 'Estoque Custo Real']], use_container_width=True)

            # --- Alertas Críticos ---
            st.divider()
            st.subheader("⚠️ Alertas de Atenção")
            
            aba1, aba2 = st.tabs(["🔥 Rupturas Iminentes (Classe A)", "❄️ Stock Parado (Micos)"])
            
            with aba1:
                if not rupturas.empty:
                    st.warning(f"Foram encontrados {len(rupturas)} itens de Curva A com menos de 7 dias de cobertura.")
                    st.table(rupturas[['Referência', 'Qtd. Estoque', 'Cobertura_Dias', 'Qtd. Média Líq.']].rename(columns={'Qtd. Estoque': 'Pares em Stock'}))
                else:
                    st.success("Nenhum item da Curva A em risco crítico de ruptura.")

            with aba2:
                if not micos.empty:
                    st.info(f"Foram encontrados {len(micos)} itens classificados como 'Mico' (baixo giro e baixa margem).")
                    st.table(micos[['Referência', 'Giro', '%Margem', 'Estoque Custo Real', 'Qtd. Estoque']].rename(columns={'Qtd. Estoque': 'Pares Parados'}))
                else:
                    st.success("Não foram detetados 'Micos' críticos no stock atual.")

            # --- Exportação ---
            st.divider()
            # Converter dataframe para Excel em memória
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Analise_Giro')
            
            st.download_button(
                label="📥 Baixar Análise Completa em Excel",
                data=buffer.getvalue(),
                file_name="resultado_analise_estoque.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro ao processar o ficheiro: {e}")
    else:
        st.info("Aguardando upload do ficheiro Excel na barra lateral.")

if __name__ == "__main__":
    main()