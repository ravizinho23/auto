import pyautogui
import time
import pandas as pd
import os
import Tratamento_Planilha as tp
import tkinter as tk
from tkinter import ttk
from playwright.sync_api import sync_playwright
import yaml

# Carrega as configurações de um arquivo YAML
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

# Função que simula a navegação por setas com base na seleção do Tkinter usando pyautogui
def navigate_with_arrows(option_index):
    time.sleep(1)  # Aguarde um pouco para garantir que o foco esteja correto
    for _ in range(option_index):
        pyautogui.press('down')  # Pressiona a seta para baixo
        time.sleep(0.5)  # Aguarde um pouco entre as pressões
    pyautogui.press('enter')  # Pressiona Enter para selecionar

async def navigate_to_option(page, option_index, select_field_selector):
    try:
        await page.click(select_field_selector)
        await page.wait_for_selector('select[id="escaninho_consult_pesq\\:unidade"] option', state='visible')  # Selecionador mais específico
        options = await page.query_selector_all('select[id="escaninho_consult_pesq\\:unidade"] option')
        if option_index < len(options):
            await options[option_index].click()
        else:
            raise ValueError(f"Índice da opção inválido: {option_index}")
    except Exception as e:
        print(f"Erro ao navegar para a opção: {e}")

def run_automation(username, password, option_value_index):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()

            # Navega para o site e faz login
            page.goto('https://www.sistemas.pa.gov.br/governodigital/public/main/index.xhtml')
            page.fill('#form_login\\:login_username', username)
            page.fill('#form_login\\:login_password', password)
            page.click('#form_login\\:button_login')
            page.wait_for_selector('#form_sistema\\:submit_area > div > div:nth-child(2) > div.SistemaGridLabel > a > p')
            page.click('#form_sistema\\:submit_area > div > div:nth-child(2) > div.SistemaGridLabel > a > p')
            time.sleep(3)  # Ajuste o tempo conforme necessário
            
            # Espera a nova página abrir
            new_page = page.wait_for_event('popup')
            new_page.wait_for_selector('#iconmenu_vert\\:panelMenuGroupProtocoloEletronico')
            new_page.click('#iconmenu_vert\\:panelMenuGroupProtocoloEletronico')
            new_page.wait_for_selector('#iconmenu_vert\\:j_id52')
            new_page.click('#iconmenu_vert\\:j_id52')
            new_page.wait_for_selector('#iconmenu_vert\\:panelMenuItemEscaninhoUnidade')
            new_page.click('#iconmenu_vert\\:panelMenuItemEscaninhoUnidade')

            time.sleep(2)
            
            # Foco no campo de seleção
            select_field_selector = config['selectors']['select_field']
            try:
                new_page.wait_for_selector(select_field_selector, state='visible', timeout=5000)  # Timeout de 5 segundos
                if not new_page.locator(select_field_selector).is_enabled():
                    raise Exception("O campo de seleção não está habilitado")
                
                # Clica no campo de seleção para focar
                new_page.click(select_field_selector)
                time.sleep(1)  # Aguarde para garantir que o campo está ativo
                
                # Navega usando as setas baseado no índice da opção selecionada no Tkinter
                navigate_with_arrows(option_value_index)

            except Exception as e:
                print(f"Erro ao interagir com o campo de seleção: {e}")

            # Extrai e armazena os dados da tabela
            df = extract_and_store_table_data(new_page)

            # Caminho do arquivo Excel
            file_path = r"C:\PROJETO SEDUC RAFAEL\optimuspae-20240919T111055Z-001\optimuspae\Novo(a) Planilha de trabalho XLS.xlsx"

            # Verifica se o arquivo já existe
            if os.path.exists(file_path):
                # Lê o arquivo existente
                existing_df = pd.read_excel(file_path)
                
                # Adiciona os novos dados abaixo dos dados existentes
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    startrow = len(existing_df) + 1
                    df.to_excel(writer, sheet_name='Sheet1', startrow=startrow, index=False, header=False)
            else:
                # Se o arquivo não existir, cria um novo arquivo
                df.to_excel(file_path, index=False)

            browser.close()
    except Exception as e:
        print(f"Erro durante a automação: {e}")

def extract_and_store_table_data(page):
    all_data = []
    previous_data = None
    consecutive_same_data_count = 0

    while True:
        try:
            # Aguarde a tabela estar presente na página
            page.wait_for_selector('#escaninho_consult_pesq\\:table tbody tr', timeout=10000)  # Timeout de 10 segundos

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
                    page.wait_for_timeout(2000)  # Ajuste o tempo de espera conforme necessário
                else:
                    break
            else:
                break
        except Exception as e:
            print(f"Erro ao extrair dados da tabela: {e}")
            # Simular a tecla Enter se a tabela não puder ser extraída
            pyautogui.press('enter')
            break

    # Convertendo os dados para um DataFrame e processando
    df = pd.DataFrame(all_data)
    df_processado = tp.processar_dataframe(df, drop_duplicates=True)
    return df_processado

def show_gui():
    def on_submit():
        username = entry_username.get()
        password = entry_password.get()
        option_value = combo_option.get()
        option_index = options.index(option_value)  # Calcula o índice da opção selecionada
        root.destroy()
        run_automation(username, password, option_index)

    root = tk.Tk()
    root.title("Dados de Login")

    tk.Label(root, text="Login:").grid(row=0, column=0, padx=10, pady=5)
    entry_username = tk.Entry(root)
    entry_username.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="Senha:").grid(row=1, column=0, padx=10, pady=5)
    entry_password = tk.Entry(root, show="*")
    entry_password.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(root, text="Unidade:").grid(row=2, column=0, padx=10, pady=5)
    options = [
        'CAPO/AN -  Aposentadoria Novo - SE01',
               'CAPO/Abono - CAPO/Abono Permanência - SE01',
               'CAPO/Formproc - CAPO - Formalidade Processual - SE01',
               'CAPO/Judicial - CAPO Aposentadoria Judicial - SE01',
               'CAPO/NãoEstavel - CAPO- Não Estáveis - SE01',
               'CAPO/PENDENTE -  Aposentadoria Pendente - SE01',
               'CAPO/PROTOCOLO - COORDENADORIA DE APOSENTADORIA  - PROTOCOLO - SE01',
               'CAPO/Triagem - Capo Triagem - SE01',
               'CCM - COORDENADORIA DE CONTROLE E MOVIMENTAÇÃO - SE01',
               'CCM/Averbação - CCM Averbação - SE01',
               'CCMCEDERAssina - Cessão e Revogação Assinatura CCM - SE01',
               'CCM CEP - CCM CEP - SE01',
               'CCM/Certidões - CCM Certidões  - SE01',
               'CCM COMUNICAÇÃO - CCM COMUNICAÇÃO - SE01',
               'CCM DEC ATES - CCM DECLARAÇÃO E ATESTADO - SE01',
               'CCM DESIGASSINA - CCM Designação e Dispensa Assinatura - SE01',
               'CCM DESIG/SUBS. - CCM Designação e Substituição - SE01',
               'CCM/Férias - CCM Férias - SE01',
               'CCM/Férias/Ass. - CCM FÉRIAS/ASSINATURA - SE01',
               'CCM JUDICIAL - CCM JUDICIAL  - SE01',
               'CCM MANUTENÇAO - CCM MANUTENÇÃO - SE01',
               'CCM/Pecúnia - CCM- Pecúnia - SE01',
               'CCM POLO - CCM POLO - SE01',
               'CCM/RED/LP - CCM/REDISTRIBUIÇÃO LICENÇA PRÊMIO - SE01',
               'CCM TRIAGEM - CCM TRIAGEM - SE01',
               'CCM TRIÊNIO - CCM TE TRIÊNIO - SE01',
               'CFOP - COORDENADORIA DE FOLHA DE PAGAMENTO - SE01',
               'CFOP/JUDICIAL - CFOP PROCESSOS JUDICIAIS  - SE01',
               'CFOP/SOBRESTADO - PROCESSOS SOBRESTADOS - SE01',
               'COR - COORDENADORIA DE ORGANIZAÇÃO DE REDE - SE01',
               'COR/CODIGO - COR CODIGO - SE01',
               'COR/JUDICIAL - COR PROCESSOS JUDICIAIS - SE01',
               'CPS - COORDENADORIA DE PLANEJAMENTO E SELEÇÃO - SE01',
               'CPS/Contratos - CPSP Contratos Temporarios - SE01',
               'CPS/Estágio - CPS Contratos Estágio - SE01',
               'CPS/JUDICIAL - CPS PROCESSOS JUDICIAIS - SE01',
               'CPS/Portarias - CPS PORTARIAS - SE01',
               'CVAS - Coordenadoria de Valorização e Assistência ao Servidor  - SE01',
               'CVAS/AVA PSIQ - CVAS Avaliação Psiquiátrica  - SE01',
               'CVAS/JUDICIAIL - CVAS PROCESSOS JUDICIAIS - SE01',
               'CVAS-LA/Assina - CVAS Licença Aprimoramento Assinatura - SE01',
               'DIFOB - DIRETORIA DE FOLHA E BENEFÍCIOS - SE01',
               'DIOP - DIRETORIA DE ORGANIZAÇÃO DE PESSOAL - SE01',
               'DIPSE DISTRATOS - DIPSE DISTRATOS TEMPORÁRIOS - SE01',
               'SAGEP - SECRETARIA ADJUNTA DE GESTÃO DE PESSOAS - SE01',
               'SAGEP CONTROLE - SAGEP ÓRGÃOS DE CONTROLE - SE01',
               'SAGEP/ESIC - SAGEP E-SIC - SE01',
               'SAGEP/JUDICIAL - SAGEP Processos Judiciais - SE01',
               'SAGEP/PAD - SAGEP - Processos Administrativo Disciplinar PAD - SE01'
    ]
    combo_option = ttk.Combobox(root, values=options)
    combo_option.grid(row=2, column=1, padx=10, pady=5)

    tk.Button(root, text="Enviar", command=on_submit).grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    show_gui()
