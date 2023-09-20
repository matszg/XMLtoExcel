import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET

#func para parsear xml
def parse_xml(xml_content):
    try:
        root = ET.fromstring(xml_content)

        namespace = {'ans': 'http://www.ans.gov.br/padroes/tiss/schemas'}

        # Encontrar e printar o nm de carteirinha
        numero_carteira = root.find('.//ans:dadosBeneficiario/ans:numeroCarteira', namespaces=namespace).text
        print("Carteirinha:", numero_carteira)

        # Encontrar e printar o nm de guia
        numero_guia = root.find('.//ans:dadosAutorizacao/ans:numeroGuiaOperadora', namespaces=namespace).text
        print("Numero da Guia:", numero_guia)
        
        # Encontrar e printar a data
        data_auto = root.find('.//ans:dadosAutorizacao/ans:dataAutorizacao', namespaces=namespace).text
        print("Data de Autorizacao:", data_auto)
        
        return numero_carteira, numero_guia, data_auto
        
    except Exception as e:
        print("Error parsing XML:", str(e))
        return None, None, None
    
#Janela de seleciona as file
def main():
    root = tk.Tk()
    root.withdraw()

    # Selecionar diretorio
    file_path = filedialog.askdirectory()

    # listagem de xml
    lista_files = [file for file in os.listdir(file_path) if file.endswith('.xml')]

    #colunas das tabelas
    coluna_xml = ["Carteirinha", "NumeroGuia", "DataAutorizacao", "Plano"]
    value_xml = []

#loop dos xml
    for xml_file in lista_files:
        xml_file_path = os.path.join(file_path, xml_file)
        with open(xml_file_path, "rb") as xmlfile:
            xml_content = xmlfile.read()
            Carteirinha, NumeroGuia, DataAutorizacao = parse_xml(xml_content)
            if Carteirinha is not None and NumeroGuia is not None and DataAutorizacao is not None:
                Plano = "PRE" if Carteirinha.startswith("0044") else "INTER"
                value_xml.append([Carteirinha, NumeroGuia, DataAutorizacao, Plano])

    pastadoexcel = os.path.join(file_path, "Fatura.xlsx")
    tabela = pd.DataFrame(columns=coluna_xml, data=value_xml)
    tabela.to_excel(pastadoexcel, index=False)
    print(f"Data saved to {pastadoexcel}")

if __name__ == "__main__":
    main()
