# Desafio João Marcos Cavacalnti de Albuquerque Neto
import openpyxl
from openpyxl import *
from openpyxl.utils import get_column_letter
wb = load_workbook(filename = 'Engenharia de Software - João Marcos Neto.xlsx')
Notas = wb.active
media = 0
for lin in range(4,Notas.max_row+1):
    if Notas['C' + str(lin)].value > (60*0.25):
        Notas['G' + str(lin)].value = "Reprovado por Falta"
    else:
        for col in range(4,7):
            char = get_column_letter(col)
            media += Notas[char + str(lin)].value/30  
            Notas['G' + str(lin)].value = media
            if (media >= 5) and (media < 7):
                if Notas['G' + str(lin)].value != "Reprovado por Falta":
                    Notas['G' + str(lin)].value = "Exame Final"
                    NotaPF = 10-media
                    Notas['H' + str(lin)].value = float("{:.1f}".format(NotaPF))
            else:
                Notas['G' + str(lin)].value = "Aprovado"
                Notas['H' + str(lin)].value = 0
        media = 0
            
        
wb.save('Desafio João Marcos Cavalcanti de Albuquerque Neto.xlsx')
