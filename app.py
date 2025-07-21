"""
1.Pegar os dados da planilha
-nome do curso
-nome do participante
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
    carga_horaria= linha[4].value #carga horaria
    dt_emissao_cert= linha[5].value #data de emissao de certificado
    
    #2.TRANSFERIR OS DADOS PARA A IMAGEM DO CCERTIFICADO
    #arrumando letras (fontes e tamanhos)
    fonte_nome = ImageFont.truetype('./src/fontes/tahomabd.ttf', 60)
    fonte_geral = ImageFont.truetype('./src/fontes/tahoma.ttf', 40)
    fonte_data = ImageFont.truetype('./src/fontes/tahoma.ttf',40)

    #abrindo a imagem do certificado
    imagem_certificado = Image.open('./src/certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem_certificado)

    #escrevendo os dados na imagem:
    #nome do participante
    desenhar.text((750, 715), nome_participante, font=fonte_nome, fill='black') 
    #nome do curso
    desenhar.text((750, 790), nome_curso, font=fonte_geral, fill='black') 
    #carga horaria 
    desenhar.text((615, 850), str(carga_horaria), font=fonte_geral, fill='black')
    #data de emissao do certificado
    desenhar.text((1180, 850), dt_emissao_cert, font=fonte_data, fill='black') 


    imagem_certificado.save(f'./src/Certificados/{indice} {nome_participante} certificado.jpg')