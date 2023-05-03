import openpyxl

letras = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'W', 'X', 'Y', 'Z'}
empresas = openpyxl.load_workbook("Planilha.xlsx")
print(empresas.sheetnames)

empresas_page = empresas['Planilha1']
with open("dados.txt", "r") as f:
    d = f.readlines()
    tamanho = len(d)
    print(d)
print(tamanho)
for i in range(tamanho):
    e = d[i].split(', ')
    print(len(e))
    try:
        for n in range(6):
            indice = letras[n]+str(i+2)
            empresas_page[indice] = e[n]
    except:
        if len(e)!=6:
            print("Empresa da linha {} está com dados faltando (Pode ser o nome ou algum dos dados).".format(i+1))
        for nn in range(1, 5):
            if " " in e[nn]:
                print("Empresa da linha {} tem espaço (" ") no meio de seus dados.")
            if chr[32:47] in e[nn] or chr[58:64] in e[nn] or chr[91:96] in e[nn] or chr[123:127] in e[nn]:
                print("Empresa da linha {} tem caractere invalido nos seus dados".format(i))
empresas.save("Planilha.xlsx")