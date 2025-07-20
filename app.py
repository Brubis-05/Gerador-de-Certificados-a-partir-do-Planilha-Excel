"""
1.Pegar os dados da planilha
-nome do curso
-nome do participante
-tipo de participação
-data do inicio
-data do final
-carga horaria
-data de emissao do certificado

2.transferir os dados para a imagem do certificado
"""

import openpyxl
from PIL import Image, ImageDraw, ImageFont

#1.PEGAR OS DADOS DA PLANILHA:
#abrindo a planilha
workbook_alunos = openpyxl.load_workbook('./src/planilha_alunos.xlsx')
sheet_alunos = workbook_alunos ['Sheet1'] #pagina do excel que vai ser utilizada

#pegando os dados da planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada cécula que contem a informação que precisamos
    nome_curso= linha[0].value #nome do curso
    nome_participante= linha[1].value #nome do participante
    tipo_de_participacao= linha[2].value #tipo de participação
    dt_inicio= linha[3].value #data do inicio
    dt_final= linha[4].value #data do final
    carga_horaria= linha[5].value #carga horaria
    dt_emissao_cert= linha[6].value #data de emissao de certificado
    
    #2.TRANSFERIR OS DADOS PARA A IMAGEM DO CCERTIFICADO
    #arrumando letras (fontes e tamanhos)
    fonte_nome = ImageFont.truetype('./src/fontes/tahomabd.ttf', 150)
    fonte_geral = ImageFont.truetype('./src/fontes/tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./src/fontes/tahoma.ttf', 55)

    #abrindo a imagem do certificado
    imagem_certificado = Image.open('./src/certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem_certificado)

    #escrevendo os dados na imagem:
    desenhar.text((1100, 750), nome_participante, font=fonte_nome, fill='black') #nome do participante
    desenhar.text((1080, 950), nome_curso, font=fonte_geral, fill='black') #nome do curso
    desenhar.text((1430, 1065), tipo_de_participacao, font=fonte_geral, fill='black') #tipo de participação
    desenhar.text((1480, 1190), str(carga_horaria), font=fonte_geral, fill='black') #carga horaria
    
    desenhar.text((750, 1780), dt_inicio, font=fonte_data, fill='black') #data do inicio
    desenhar.text((750, 1930), dt_final, font=fonte_data, fill='black') #data do final
    
    desenhar.text((2220, 1930), dt_emissao_cert, font=fonte_data, fill='black') #data de emissao do certificado


    imagem_certificado.save(f'./src/Certificados/{indice} {nome_participante} certificado.jpg')