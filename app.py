"""

# Pegar os dados da planilha
Tipo nome do curso, nome partivipante, tipo de participação, data do inicio, data final,
carga horaia, data de emissão do certificado e as assinaturas do Gesto geral, do coodenador e dos aluno.



# Transferir para a Imagem do certificado
"""

import openpyxl
from PIL import Image, ImageDraw, ImageFont, Image


#Abrir a planilha
workbook_alunos = openpyxl = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    #Cada célula que contem a info que precisamos
    nome_curso = linha[0].value # Nome do curso
    nome_participante = linha[1].value # Nome do participante
    tipo_participacao = linha[2].value # Tipo de participação
    data_inicio = linha[3].value # Data inicio
    data_final = linha[4].value # Data Final
    carga_hora = linha[5].value # Carga horaria
    data_emissao = linha[6].value # Data da emissão
    

    # Transferir os dados da planilha para a imagem do certificado
    # Definindo a fonte a ser usado
    font_nome = ImageFont.truetype('./tahomaBD.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_inicio_fim = ImageFont.truetype('./tahoma.ttf', 50)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # Definindo coodenadas para escrever as info
    # Dica: Usar o paint com a imagem aberta e posicionar o mouse para ter um valor aproximado.
    desenhar.text((1016,825), nome_participante, fill = 'black', font = font_nome)
    desenhar.text((1069,955), nome_curso, fill = 'black', font = fonte_geral)
    desenhar.text((1435,1070), tipo_participacao, fill = 'black', font = fonte_geral)
    desenhar.text((1485,1190), str(carga_hora)+" h", fill = 'black', font = fonte_geral)
    desenhar.text((755,1780), data_inicio, fill = 'black', font = fonte_inicio_fim)
    desenhar.text((755,1930), data_final, fill = 'black', font = fonte_inicio_fim)
    desenhar.text((2225,1930), data_emissao, fill = 'black', font = fonte_inicio_fim)


    image.save(f'./{indice} {nome_participante} certificados.png')

