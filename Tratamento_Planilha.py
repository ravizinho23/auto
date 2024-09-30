import pandas as pd
import os
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

def extract_and_store_table_data(page):
    all_data = []
    previous_data = None
    consecutive_same_data_count = 0

    while True:
        try:
            # Aguarde a tabela estar presente na página
            page.wait_for_selector('#escaninho_consult_pesq\\:table tbody tr')

            # Extrair os dados da tabela na página atual
            table_rows = page.locator('#escaninho_consult_pesq\\:table tbody tr').all()
            table_data = []
            for row in table_rows:
                columns = row.locator('td').all()
                row_data = [column.evaluate('(element) => element.textContent') for column in columns]
                table_data.append(row_data)
            if table_data == previous_data:
                consecutive_same_data_count += 1
            else:
                consecutive_same_data_count = 0

            # Verifique se a próxima página é igual à anterior
            if consecutive_same_data_count >= 3:
                break
            previous_data = table_data
            all_data.extend(table_data)
            # Verificar se há próxima página e navegar, se disponível
            next_page_buttons = page.locator('.rich-datascr-button[onclick*="\'page\': \'next\'"]')
            if next_page_buttons.count() > 0:
                next_page_button = next_page_buttons.nth(-1)
                if not next_page_button.is_disabled():
                    next_page_button.click()
                    # Aguarde um tempo para o carregamento da próxima página
                    page.wait_for_timeout(2000)
                else:
                    break
            else:
                break
        except Exception as e:
            print(f"Erro ao extrair dados da tabela: {e}")
            break

    # Convertendo os dados para um DataFrame e processando
    df = pd.DataFrame(all_data)
    df_processado = processar_dataframe(df, drop_duplicates=True)
    return df_processado

def processar_dataframe(df, drop_duplicates=False):
    # Remove duplicatas se necessário
    if drop_duplicates:
        df = df.drop_duplicates()

    # Remove colunas indesejadas, se necessário
    col_to_remove = "rich-datascr-button-dsbld rich-datascr-button"  # Altere para o nome real da coluna se necessário
    if col_to_remove in df.columns:
        df = df.drop(columns=[col_to_remove])

    df = df[~df.apply(lambda row: row.astype(str).str.contains('««').any(), axis=1)]

    return df