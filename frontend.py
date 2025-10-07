import streamlit as st
from datetime import datetime
from backend import processar_arquivo

# Configurar a página
st.set_page_config(
    page_title="Processador de Report de Planejamento",
    page_icon="📋",
    layout="wide"
)

# Interface principal
st.title("🎨 Processador de Report de Planejamento")
st.markdown("---")

# Instruções de uso no topo (acima do upload)
with st.expander("ℹ️ INSTRUÇÕES DE USO", expanded=True):
    st.markdown("""
    ### 📋 Como usar:
    1. **Carregue o arquivo Excel** usando o botão abaixo
    2. **Clique em 'Processar e Formatar Arquivo'** para iniciar
    3. **Aguarde** o processamento e formatação
    4. **Baixe o arquivo formatado** quando estiver pronto
    
    ### 🎨 Esquema de Cores Aplicado:
    
    **Cabeçalhos:**
    - Fundo cinza (#8F8D8D) com texto preto
    
    **Linhas de Dados:**
    - 🔵 **Azul Claro (#4AAFBD)**: Itens onde PROCUREMENT_KEY = "E"
    - 🔷 **Azul Escuro (#061569)**: Itens onde DEMAND começa com "R" (texto branco)
    - 🟢 **Verde (#38F58A)**: Itens onde Status = "In Stock"
    
    ### 🔤 Fonte Aplicada:
    - **Aptos Narrow tamanho 11** em toda a planilha
    - **Negrito** nos títulos e cabeçalhos da aba Overview
    """)

st.markdown("---")

uploaded_file = st.file_uploader("Carregue seu arquivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
    
    # Botão de processar em cinza (secondary)
    if st.button("🎨 PROCESSAR E FORMATAR ARQUIVO", type="secondary"):
        with st.spinner("Processando e formatando arquivo... Aguarde"):
            try:
                output, total_bi, total_rm, total_po, total_stock = processar_arquivo(uploaded_file)
                
                st.success("✅ Processamento e formatação concluídos com sucesso!")
                
                # Botão de download em verde (primary)
                data_hoje = datetime.now().strftime("%Y%m%d")
                nome_arquivo = f"{data_hoje} - Rotina de planejamento.xlsx"
                
                st.download_button(
                    label="📥 BAIXAR ARQUIVO FORMATADO",
                    data=output,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"  # Verde
                )
                
                # Resumo
                st.subheader("📋 Resumo do Processamento")
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("Overview", "Análises")
                col2.metric("Relatório BI", f"{total_bi} linhas")
                col3.metric("Aba RM", f"{total_rm} linhas")
                col4.metric("Aba PO", f"{total_po} linhas")
                col5.metric("Aba Stock", f"{total_stock} linhas")
                
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")
else:
    st.info("👆 Por favor, carregue um arquivo Excel para começar.")

# Estrutura das abas
with st.expander("📋 Estrutura das Abas"):
    st.markdown("""
    **Overview**: Análises consolidadas com formatação simplificada
    **Relatório BI**: Dados completos do arquivo original
    **RM**: Itens com STATUS_STYPE = 'PurRequist' 
    **PO**: Itens com STATUS_STYPE = 'POConfirm' ou 'POCreated' 
    **Stock**: Itens com Status = 'In Stock' 
    """)

# Footer com versão e criador
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        <p>Versão 1.0 | Desenvolvido por Tiago Rocha</p>
    </div>
    """, 
    unsafe_allow_html=True
)