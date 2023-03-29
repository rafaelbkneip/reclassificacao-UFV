import requests
import xlsxwriter  
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from datetime import date

options = Options()
options.add_experimental_option("detach", True)

navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador.get("https://www2.pse.ufv.br/wp-content/uploads/2023/03/sisu2023-chamada-2.htm")

#Quadro inicial com os cursos que apresentam novos alunos, separados por campus
lista_cursos = navegador.find_elements(By.ID, 'indice')
cursos = lista_cursos[0].text.split("\n")

#Inicia lista para deletar informações que não importantes
deletar = []
cont = 0

#Deleta primeira variável da lista, título do quadro
cursos.pop(0)

#Define as posições da lista que devem ser excluídos para que 'cursos' seja apenas uma lista de cursos
for i in range(len(cursos)):
    if(cursos[i].split(" ")[0] == 'CAMPUS'):
        deletar.append(i)

#Apaga da lista as variáveis nas posições definidas anteriormente
for r in range(len(deletar)):
  cursos.pop( deletar[r] - cont )
  cont= cont +1

print(cursos)

n_cursos = (len(cursos))
print(n_cursos)

#Define lista para as informações
inscricao = []
nome = []
pontos = []
modalidade_inscricao = []
modalidade_convocacao = []
curso_aluno = []

#Os alunos aprovados são mostrados a partir do terceiro quadro
for j in range(3, n_cursos+3):
    cont2 = 2
    curso = (cursos[j-3])

    while(True):
        #Salva na lista as informações dos alunos até o fim do quadro para aquele curso
        try:
            inscricao.append(navegador.find_element(By.XPATH, '/html/body/table[' + str(j) + ']/tbody/tr[' + str(cont2) + ']/td[1]').text)
            nome.append(navegador.find_element(By.XPATH, '/html/body/table[' + str(j) + ']/tbody/tr[' + str(cont2) + ']/td[2]').text)
            curso_aluno.append(curso)
            pontos.append(navegador.find_element(By.XPATH, '/html/body/table[' + str(j) + ']/tbody/tr[' + str(cont2) + ']/td[3]').text)
            modalidade_inscricao.append(navegador.find_element(By.XPATH, '/html/body/table[' + str(j) + ']/tbody/tr[' + str(cont2) + ']/td[4]').text)
            modalidade_convocacao.append(navegador.find_element(By.XPATH, '/html/body/table[' + str(j) + ']/tbody/tr[' + str(cont2) + ']/td[4]').text)
            
            print(inscricao[-1])
            print(nome[-1])
            print(pontos[-1])
            print(modalidade_inscricao[-1])
            print(modalidade_convocacao[-1])

            cont2 = cont2 + 1
            
        except:
            break; 

   
#Definir o caminho para salvar o arquivo .xlsx      
book = xlsxwriter.Workbook('')     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Inscrição')
sheet.write(0, 1, 'Nome')
sheet.write(0, 2, 'Aluno')
sheet.write(0, 3, 'Pontos')
sheet.write(0, 4, 'Modalidade inscrição')
sheet.write(0, 5, 'Modalidade convocação')

#Todas as listas possuem o mesmo número de elementos
for i in range(len(inscricao)):
    sheet.write(i+1, 0, inscricao[i])
    sheet.write(i+1, 1, nome[i])
    sheet.write(i+1, 2, curso_aluno[i])
    sheet.write(i+1, 3, pontos[i])
    sheet.write(i+1, 4, modalidade_inscricao[i])
    sheet.write(i+1, 5, modalidade_convocacao[i])
    
book.close()

