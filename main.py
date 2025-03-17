import PySimpleGUI as sg
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import win32com.client as win32

pagina = 1
choice = '1'
link = ""
ja_aberto = False

def misturar_arrays(*arrays):
    # Verifique o comprimento mínimo das arrays
    min_length = min(len(arr) for arr in arrays)

    # Inicialize a lista resultante
    resultado = []

    # Itere através dos índices até o comprimento mínimo
    for i in range(min_length):
        # Crie uma nova array com os elementos do índice i de todas as arrays
        nova_array = [arr[i] for arr in arrays]
        # Adicione a nova array ao resultado
        resultado.append(nova_array)

    return resultado


def obter_posicao(string):
    mapeamento = {
        'Enem_2021': '1',
        'Enem_2020': '2',
        'Enem_2019': '3',
        'Enem_port': '4',
        'Enem_mat': '5',
        'Enem_fis': '6',
        'port': '7',
        'quim': '8',
        'lit': '9',
        'ing': '10',
        'bio': '11',
        'hist': '12',
        'geo': '13',
        'filo': '14'
    }
    return mapeamento.get(string, 15)

def divide_lista(lista, tamanho):
   return [lista[i:i+tamanho] for i in range(0, len(lista), tamanho)]

def dividir_array_por_letra(array):
    resultado = {'A': [], 'B': [], 'C': [], 'D': [], 'E e espaço': []}

    for elemento in array:
        if elemento.startswith('A'):
            resultado['A'].append(elemento)
        elif elemento.startswith('B'):
            resultado['B'].append(elemento)
        elif elemento.startswith('C'):
            resultado['C'].append(elemento)
        elif elemento.startswith('D'):
            resultado['D'].append(elemento)
        else:
            resultado['E e espaço'].append(elemento)

    return resultado

# Defina o layout da interface
layout = [
    [sg.Text("SeemBotPy.exe", font=("Helvetica", 20))],

    [sg.Text("Selecione um diretório de arquivo:")],
    [sg.InputText(key="diretorio"), sg.FolderBrowse()],
    [sg.Text("Escolha uma opção:")],
    [sg.Listbox(values=['Enem_2021', 'Enem_2020', 'Enem_2019', 'Enem_port', 'Enem_mat', 'Enem_fis', 'Português' ,'Química', 'Literatura', 'Inglês', 'Biologia', 'História', 'Geografia', 'Filosofia' ], select_mode=sg.LISTBOX_SELECT_MODE_SINGLE, size=(20, 3), key="opcao")],
    [sg.Button("Abrir") ,sg.Button("Processar"), sg.Button("Cancelar"), sg.Checkbox(text='varredura de páginas', key='varrer')]
]

# Crie a janela da interface
janela = sg.Window("SeemBotPy.exe", layout)

while True:
    evento, valores = janela.read()
    diretorio_selecionado = valores["diretorio"]
    opcao_selecionada = valores["opcao"][0]
    choice = obter_posicao(valores["opcao"][0])
    if evento == sg.WIN_CLOSED or evento == "Cancelar":
        break
    if evento == "Abrir":
        if ja_aberto == True:
            driver.quit()
        ja_aberto = True


        if choice in ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14'):

            if choice == '1':
                link = f"https://estudeprisma.com/questoes/s/enem/2021/iy?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # enem 2021
                nome = 'enem_2021'

            elif choice == '2':
                link = f"https://estudeprisma.com/questoes/s/enem/2020/iy?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # enem 2020
                nome = 'enem_2020'

            elif choice == '3':
                link = f"https://estudeprisma.com/questoes/s/enem/2019/iy?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # enem 2019
                nome = 'enem_2019'

            elif choice == '4':
                link = f"https://estudeprisma.com/questoes/s/portugues/enem/di?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # port enem
                nome = 'enem_port'

            elif choice == '5':
                link = f"https://estudeprisma.com/questoes/s/matematica/enem/di?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # mat enem
                nome = 'enem_mat'

            elif choice == '6':
                link = f"https://estudeprisma.com/questoes/s/fisica/enem/di?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # fis enem
                nome = 'enem_fis'

            elif choice == '7':
                link = f"https://estudeprisma.com/questoes/s/portugues/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # port
                nome = 'port'

            elif choice == '8':
                link = f"https://estudeprisma.com/questoes/s/quimica/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # quim
                nome = 'quim'

            elif choice == '9':
                link = f"https://estudeprisma.com/questoes/s/literatura/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # lit
                nome = 'lit'

            elif choice == '10':
                link = f"https://estudeprisma.com/questoes/s/ingles/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # ing
                nome = 'ing'

            elif choice == '11':
                link = f"https://estudeprisma.com/questoes/s/biologia/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # bio
                nome = 'bio'

            elif choice == '12':
                link = f"https://estudeprisma.com/questoes/s/historia/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # hist
                nome = 'hist'

            elif choice == '13':
                link = f"https://estudeprisma.com/questoes/s/geografia/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # geo
                nome = 'geo'

            elif choice == '14':
                link = f"https://estudeprisma.com/questoes/s/filosofia/d?hasTeacherComment=true&page={pagina}&trkContext=all%20questions%20tab"
                # filo
                nome = 'filo'

        else:
            pass

        # Abriu o driver
        driver = webdriver.Chrome()
        driver.get(link)

        # espera perfeita
        WebDriverWait(driver, 0).until(
            EC.presence_of_element_located((By.TAG_NAME, "body")) and
            EC.visibility_of_element_located((By.TAG_NAME, "body"))
        )


    if evento == "Processar":
        buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Gabarito comentado')]")

        for button in buttons:
            time.sleep(0.1)
            button.click()

        # espera perfeita
        WebDriverWait(driver, 0).until(
            EC.presence_of_element_located((By.TAG_NAME, "body")) and
            EC.visibility_of_element_located((By.TAG_NAME, "body"))
        )

        # Obter o código HTML da página
        html = driver.page_source

        # Criar um objeto BeautifulSoup
        site = BeautifulSoup(html, "html.parser")
        soup = BeautifulSoup(html, "html.parser")

        # caça os textos das questões e armazena em uma array
        questao = site.find_all("div", class_="MuiBox-root css-1m6borv")

        textos = []

        # varre o codigo pegando seu texto

        for div in questao:
            frase = div.get_text()
            frase = frase.strip().replace('\n', ' ')  # polimento do texto
            textos.append(frase)

        alternativaL = site.find_all("span", class_="MuiTypography-root MuiTypography-body2 css-u7juvv")
        alternativaT = site.find_all("div", class_="MuiTypography-root MuiTypography-body2 css-12l80v4")

        respostas = []
        alternativaLFL = []
        alternativaTFL = []

        for div in alternativaL:
            alternativaLF = div.get_text()
            alternativaLFL.append(alternativaLF)

        for span in alternativaT:
            alternativaTF = span.get_text()
            alternativaTFL.append(alternativaTF)

        for i in range(len(alternativaLFL)):
            respostas.append(alternativaLFL[i] + ") " + alternativaTFL[i])

        # deixar todas com 5

        respostas_finais = []

        for i in range(len(respostas) - 1):
            if respostas[i][0] == "D" and respostas[i + 1][0] == "A":
                respostas_finais.append(respostas[i])
                respostas_finais.append("")  # Add "X" when D is followed by A
            else:
                respostas_finais.append(respostas[i])

        # Add the last item from 'respostas' to 'respostas_finais'
        #respostas_finais.append(respostas[-1])

        # Encontrar todas as divs com a classe "MuiBox-root css-ek4z02"
        divs = soup.find_all("div", class_="MuiBox-root css-ek4z02")

        # Crie uma lista vazia para armazenar os textos
        # Your existing code remains the same

        ALT = []

        # Itere sobre os divs encontrados e adicione o texto à lista
        for div in divs:
            texto = div.text
            ALT.append(texto)

        frases_sem_n = []

        for frase in ALT:
            frase_sem_n = frase.replace("\n", "")
            penultima_letra = frase_sem_n[-2]
            frases_sem_n.append(penultima_letra)

        for i in range(len(frases_sem_n)):
            frase = frases_sem_n[i]
            nova_frase = ""
            for letra in frase:
                if letra not in ['A', 'B', 'C', 'D', 'E']:
                    nova_frase += '0'
                else:
                    nova_frase += letra
            frases_sem_n[i] = nova_frase

        for i in range(len(frases_sem_n)):
            frase = frases_sem_n[i]
            nova_frase = ""
            for letra in frase:
                if letra == "A":
                    nova_frase += "1"
                elif letra == "B":
                    nova_frase += "2"
                elif letra == "C":
                    nova_frase += "3"
                elif letra == "D":
                    nova_frase += "4"
                elif letra == "E":
                    nova_frase += "5"
                else:
                    nova_frase += letra

            frases_sem_n[i] = nova_frase



        colunas = ['Type', 'Question', 'Mark','Correct Ans', 'Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5',]


        resultado = dividir_array_por_letra(respostas_finais)
        array_A = resultado['A']
        array_B = resultado['B']
        array_C = resultado['C']
        array_D = resultado['D']
        array_E_espaco = resultado['E e espaço']
        data = []

        m = []
        Mark = []
        for i in array_A:
            m.append('M')
            Mark.append('1')

        data = misturar_arrays(m, textos, Mark, frases_sem_n, array_A, array_B, array_C, array_D, array_E_espaco)
        nova_data = []

        for array in data:
            # Verifique se o quarto item (índice 3) é igual a 0
            if array[3] != '0':
                # Se não for igual a 0, adicione o array à nova lista
                nova_data.append(array)
        print (len(nova_data))

        tabela = pd.DataFrame(columns=colunas, data=nova_data)
        tabela.to_excel(f"questões_{nome}.xlsx", index=False)


        # Realize o processamento aqui, incluindo o uso das variáveis 'diretorio_selecionado' e 'opcao_selecionada'
        sg.popup("Processamento concluído!", "Diretório selecionado: " + diretorio_selecionado, "Opção selecionada: " + opcao_selecionada)

# Feche a janela quando terminar
janela.close()

#organizar formatação
#fazer programa de varredura de pagina
#integração com chat gpt
