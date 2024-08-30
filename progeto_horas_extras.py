import os
import pandas as pd
import streamlit as st
from datetime import datetime
import platform

# Função para a Página 1: Filtrando Horas Extras
def pagina_filtragem_horas_extras():
    st.title("Filtrando Horas Extras")

    uploaded_file = st.file_uploader("Escolha o arquivo a ser transformado", type=["csv", "xlsx", "txt"])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith(".xlsx"):
                df = pd.read_excel(uploaded_file)
            elif uploaded_file.name.endswith(".txt"):
                df = pd.read_csv(uploaded_file, sep='\t')
            
            # Converte a coluna de horas extras para datetime
            df["Horas Extr."] = pd.to_datetime(df["Horas Extr."], format='%H:%M').dt.time
            
            # Filtra as colunas desejadas
            filtro = df[["Colaborador", "Função", "Data", "Horas Extr.", "CPF"]]
            
            # Filtros adicionais
            filtro2 = filtro[(filtro["Horas Extr."] > datetime.strptime("02:00", "%H:%M").time()) & (filtro["Horas Extr."] < datetime.strptime("02:59", "%H:%M").time())]
            filtro3 = filtro[(filtro["Horas Extr."] > datetime.strptime("03:00", "%H:%M").time()) & (filtro["Horas Extr."] < datetime.strptime("03:59", "%H:%M").time())]
            filtro4 = filtro[filtro["Horas Extr."] > datetime.strptime("04:00", "%H:%M").time()]

            # Exibe a tabela filtrada padrão
            st.dataframe(filtro, use_container_width=True)
            
            # Escolha do filtro para exportação
            opcao_exportacao = st.radio(
                "Escolha qual filtro deseja exportar:",
                ("Filtro padrão", "Filtro 2: Horas entre 02:00 e 02:59", "Filtro 3: Horas entre 03:00 e 03:59", "Filtro 4: Horas acima de 04:00")
            )
            
            # Define o DataFrame a ser exportado
            if opcao_exportacao == "Filtro padrão":
                df_export = filtro
                nome_arquivo_base = "Tabela_Filtrada"
            elif opcao_exportacao == "Filtro 2: Horas entre 02:00 e 02:59":
                df_export = filtro2
                nome_arquivo_base = "Tabela_Filtrada_2"
            elif opcao_exportacao == "Filtro 3: Horas entre 03:00 e 03:59":
                df_export = filtro3
                nome_arquivo_base = "Tabela_Filtrada_3"
            else:
                df_export = filtro4
                nome_arquivo_base = "Tabela_Filtrada_4"

            # Converte o DataFrame selecionado para CSV
            csv = df_export.to_csv(index=False, sep=",", encoding="UTF-8")
            
            # Define o nome do arquivo com a data atual
            data_atual = datetime.now().strftime("%d-%m-%Y")
            nome_arquivo = f"{nome_arquivo_base}_{data_atual}.csv"
            
            # Botão de download
            st.download_button(
                label="Baixar o Arquivo Filtrado",
                data=csv,
                file_name=nome_arquivo,
                mime='text/csv'
            )
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
    else:
        st.warning("Por favor, carregue um arquivo para visualizar os dados.")

# Função para combinar arquivos CSV
def combinar_csv(uploaded_files, nome_arquivo):
    df_list = []
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_csv(uploaded_file)
            df_list.append(df)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {uploaded_file.name}: {e}")

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        csv_combined = combined_df.to_csv(index=False, sep=",", encoding="UTF-8")
        return csv_combined
    else:
        st.warning("Nenhum arquivo CSV válido foi processado.")
        return None

# Função para baixar anexos CSV do Outlook (Somente para Windows)
def baixar_anexos_csv_outlook(pasta_destino):
    if platform.system() == "Windows":
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)  # A pasta 6 é a Caixa de Entrada
            messages = inbox.Items
            csv_files = []
            
            for message in messages:
                attachments = message.Attachments
                for attachment in attachments:
                    if attachment.FileName.endswith('.csv'):
                        # Salva o anexo na pasta de destino
                        attachment.SaveAsFile(os.path.join(pasta_destino, attachment.FileName))
                        csv_files.append(attachment.FileName)

            return csv_files

        except Exception as e:
            st.error(f"Erro ao acessar o Outlook: {e}")
            return []
    else:
        st.error("Este recurso está disponível apenas para Windows.")
        return []

# Função principal
def main():
    st.sidebar.title("Navegação")
    pagina = st.sidebar.radio("Escolha a página:", ["Filtrando Horas Extras", "Juntar Arquivos CSV", "Baixar Arquivos do Outlook"])

    if pagina == "Filtrando Horas Extras":
        pagina_filtragem_horas_extras()
    elif pagina == "Juntar Arquivos CSV":
        st.title("Juntar Arquivos CSV")

        uploaded_files = st.file_uploader("Escolha os arquivos CSV para combinar", type=["csv"], accept_multiple_files=True)
        nome_arquivo = st.text_input('Nome do arquivo CSV combinado (ex: combinado.csv)', 'combinado.csv')

        if st.button("Combinar Arquivos") and uploaded_files:
            csv_combined = combinar_csv(uploaded_files, nome_arquivo)
            if csv_combined:
                st.download_button(
                    label="Baixar Arquivo Combinado",
                    data=csv_combined,
                    file_name=nome_arquivo,
                    mime='text/csv'
                )
    elif pagina == "Baixar Arquivos do Outlook":
        st.title("Baixar Arquivos do Outlook")

        pasta_destino = st.text_input("Escolha a pasta de destino para salvar os anexos", "")
        pasta_destino = os.path.expanduser(pasta_destino)  # Expand user paths like "~"

        if st.button("Baixar Anexos CSV"):
            if pasta_destino:
                try:
                    csv_files = baixar_anexos_csv_outlook(pasta_destino)
                    if csv_files:
                        st.success(f"Anexos CSV baixados para: {pasta_destino}")
                    else:
                        st.warning("Nenhum anexo CSV encontrado.")
                except Exception as e:
                    st.error(f"Erro ao baixar anexos: {e}")
            else:
                st.error("Por favor, insira um caminho válido para a pasta de destino.")

if __name__ == "__main__":
    main()
