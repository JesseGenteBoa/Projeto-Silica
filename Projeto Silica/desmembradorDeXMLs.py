from tkinter.filedialog import askopenfilenames
from tkinter import messagebox, Tk
import xmltodict
import copy
import shutil
import sys
import os
 
 
root = Tk()
root.withdraw()
 
 
try:
    xml_a_desmembrar = askopenfilenames(title="Selecione o arquivo XML que deseja desmembrar.", filetypes=[("XML Files", "*.xml")])
except:
    messagebox.showerror("Erro!", "Insira um ou mais arquivos XML válidos. A automação aceita somente arquivos com extenção '.xml'.")
    root.update()
    raise Exception ("Insira um ou mais arquivos XML válidos. A automação aceita somente arquivos com extenção '.xml'.")
 
 
def ler_xml(arquivo):
    try:
        with open(arquivo) as fd:
            doc = xmltodict.parse(fd.read())
    except UnicodeDecodeError:
        with open(arquivo, encoding='ISO-8859-1') as fd:
            doc = xmltodict.parse(fd.read())
    except:
        with open(arquivo, encoding='ISO-8859-1') as fd:
            doc = xmltodict.parse(fd.read(), attr_prefix="@", cdata_key="#text")
    return doc
 

if not xml_a_desmembrar:
    sys.exit()
    
try:
    os.mkdir("XMLs")
except FileExistsError:
    shutil.rmtree("XMLs")
    os.mkdir("XMLs")
except PermissionError:
    messagebox.showerror("Já existe uma pasta chamada 'XMLs' neste diretório. Exclua a pasta e tente executar novamente a automação.")
    root.update()
    raise Exception ("Já existe uma pasta chamada 'XMLs' neste diretório. Exclua a pasta e tente executar novamente a automação.")


for arq_xml in xml_a_desmembrar:

    xml = ler_xml(arq_xml)
    conjunto_xmls = xml["Workbook"]["Worksheet"]["Table"]["Row"]
    row_final = conjunto_xmls[-1]
    xml_copy = copy.deepcopy(xml)
    del xml_copy["Workbook"]["Worksheet"]["Table"]
    del xml_copy["Workbook"]["Worksheet"]["WorksheetOptions"]
    final_xml = xml["Workbook"]["Worksheet"]["WorksheetOptions"]
    
    primeiro_row = conjunto_xmls[0]
    table = {'@ss:ExpandedColumnCount': '73', '@ss:ExpandedRowCount': '3', '@x:FullColumns': '1', '@x:FullRows': '1', '@ss:DefaultRowHeight': '14.4'}


    aux = 0
    if len(conjunto_xmls) > 3:
        for xml in conjunto_xmls:
            if aux == 0:
                aux+=1
                continue

            aux+=1
            if aux == len(conjunto_xmls):
                break

            valor = xml["Cell"][21]["Data"]["#text"]
            valor2 = xml["Cell"][25]["Data"]["#text"]

            row_final["Cell"][1]["Data"]["#text"] = "1"
            row_final["Cell"][2]["Data"]["#text"] = valor
            row_final["Cell"][4]["Data"]["#text"] = valor2

        
            numero_nf = xml["Cell"][1]["Data"]["#text"]
            caminho_xml = "XMLs\\NF " + numero_nf + ".xml"

            xml_completo = xml_copy

            xml_completo["Workbook"]["Worksheet"]["Table"] = table
            xml_completo["Workbook"]["Worksheet"]["Table"]["Row"] = [primeiro_row, xml, row_final]
            xml_completo["Workbook"]["Worksheet"]["WorksheetOptions"] = final_xml

            nfs = xmltodict.unparse(xml_completo, pretty=True, indent="  ", short_empty_elements=True)

            parts = nfs.split('\n', 1)
            nfs_completo = parts[0] + '\n<?mso-application progid="Excel.Sheet"?>\n' + parts[1]

            with open(caminho_xml, 'w', encoding='utf-8') as arquivo_xml:
                arquivo_xml.write(nfs_completo)

        
    else:
        messagebox.showinfo("Aviso!", "O arquivo XML inserido refere-se a somente uma NF. Sendo assim, não é necessário utilizar a automação para esse caso.")
        root.update()
       
       
messagebox.showinfo("Sucesso!", "Processo concluído. Os arquivos XML foram criados com sucesso!")
root.update()