import sys
import getopt
import pandas as pd
import re
import string
from collections import Counter

############### PARAMETROS #####################
top = 30 #número de palavras que vai pegar no ranking de palavras que mais aparecem 
################################################

#### COMPANHIAS FILTROS #####################################################################

#AZ	Azul
#CI	Azul
#HA	HORIZON AIR
#JA	JAL-0190
#KM	KLM-0170
#KL	KLM-0190
#WD	Skywest-Delta
#SY	Fuji Dreams
#BL	Belavia-0170
#JL	JAL-0170
#AK	Skywest-Alaska
#BV	Belavia-0190
#NC	AIRLINK
#AJ	AZAL
#UA	United Airlines-Mesa
#TJ	Tianjin Airlines
#MA	American Airlines
#RT	United Airlines-Republic
#SY	Fuji Dreams
#SK	Skywest
#CG	COLORFUL GUIZHOU AIRLINES
#IA	ARKIA
#AA	Austral
#DS	Skywest-Delta
#UA	United Airlines-Mesa
#RT	United Airlines-Republic
#AJ	AZAL
#CG	COLORFUL GUIZHOU AIRLINES
#IA	ARKIA
#WI	WIDEROE

#Colocar o código da companhia aerea
#EX: Para pegar companhia aerea "HORIZON AIR" colocar no prompt "HA", conforme tabela acima
#EX: Para pegar mais de uma companhia aerea colocar no prompt "código1|código2|...|códigoN"

test_text = input ("Código da companhia: (Obs: Digitar 'TUDO' para todas) ")

###########################################################################################

df = pd.read_excel('Dados.xlsx', sheet_name='Dados') #pega dados brutos do excel


############################### APLICA FILTRO POR COMPANHIA AÉREA #########################
if test_text != "TUDO":
    df = df[df["STR_CLIENTE"].astype(str).str.contains(test_text, na=False)]
###########################################################################################

df2 = pd.read_excel('Dados.xlsx', sheet_name='Stopwords')# dataframe de stopwords

stopw = df2["Stopwords"].to_numpy() #passa de df para lista de stopwords

df["Item Desc."] = df["Item Desc."].map(str.strip) #arranca espaços no final e começo da frase
       
test = pd.DataFrame(df)

test = test["Item Desc."].str.lower().str.split() #pega a coluna que interessa e passa para minusculo e separa

######################################## Limpeza das palavras #####################################

jona = test.apply(lambda x: [item for item in x if item not in stopw])#Arranca as stopwords

a =  jona.apply(' '.join)

a = a.replace(['1','2','3','4','5','6','7','8','9','0'],'', regex=True)#dibuia os número
  
a = a.str.replace('[{}]'.format(string.punctuation), '')#dibuia as pontuação tudo

a = a.replace([' a ',' b ',' c ',' d ',' e ',' f '],'', regex=True)#dibuia as letra sozinha
a = a.replace([' g ',' h ',' i ',' j ','k ',' l '],'', regex=True)
a = a.replace([' m ',' n ',' o ',' p ',' q ',' r '],'', regex=True)
a = a.replace([' s ',' t ',' u ',' w ',' y ',' x ',' z '],'', regex=True)

###################################################################################################

top_words = Counter(" ".join(a).split()).most_common(top)#top 'top' palavras

new = a.str.split(" ", expand = True)#divide uma palavra por coluna

jon = pd.DataFrame(top_words)#dataframe com palavras mais frequentes

##### cria um vetor para servir de indice, contendo o numero de cada frase
m = []
for i in range(len(new)):
    m.append(int(i))

##################################### LOOP da esquerda para direita ##############################
fa = pd.DataFrame

cleito = new
        
cleito.insert(0, 'OH', m, allow_duplicates=False)

for i in range(len(jon)) :

    column = 0
    while column < (len(cleito.columns)-2):
      
       
        contain_values = cleito[cleito[column].astype(str).str.contains(jon[0][i])]
        
        f = contain_values['OH'].map(str) + " " + contain_values[column].map(str) + " " + contain_values[column + 1].map(str)

        f = f.str.split(" ", n = 1, expand = True)

        if column != 0: #concatena só se tiver coisa pa concatenar
            frames = [f, fa]
            result = pd.concat(frames)
            fa = result
        else:
            if i == 0:
                fa = f
            else:
                fa = result

        column = column + 1

        
###################################################################################################

##################################### LOOP 2 da direita para esquerda #############################


for i in range(len(jon)) :

    column = (len(cleito.columns)-2)
    
    while column >= 1:

        contain_values = cleito[cleito[column].astype(str).str.contains(jon[0][i])]

        f = contain_values['OH'].map(str) + " " + contain_values[column - 1].map(str) + " " + contain_values[column].map(str)

        f = f.str.split(" ", n = 1, expand = True)

        if column != (len(cleito.columns)-1): #concatena só se tiver coisa pa concatenar
            frames = [f, fa]
            result2 = pd.concat(frames)
            fa = result2
        else:
            if i == 0:
                fa = f
            else:
                fa = result2

        column = column - 1

            
###################################################################################################

frames = [result, result2]

result_final = pd.concat(frames) # junta os dois resultados (esq para dir e dir para esq)

result_final= result_final.drop_duplicates(subset=[0,1]) #remove as duplicadas para resultados iguais e na mesma frase

result_final = result_final[1]

result_final = result_final[~result_final.str.contains("None")] #retira da lista resultadso contendo "none"

result_final = result_final.map(str.strip) #retira espaços no comeco e no fim da frase

result_final = result_final[result_final.str.count(' ') + 1 > 1] # fica apenas com resultados que possuem duas palavras

rejeitar = pd.read_excel('Database.xlsx', sheet_name='Regex') #pega memoria de resultados espúrios


############################ LOOP para perguntar se o resultado ta massa ############################
sair = 0
while sair == 0:

    #loop que retira palavras espúrias guardados na memória
    for i in range(len(rejeitar)) :
        result_final = result_final[~result_final.str.contains(rejeitar[0][i])]

    # pega o top 10 resultado de par de palavras que aparecem
    top_words_final = Counter(result_final).most_common(10)

    AEE = pd.DataFrame(top_words_final)

    AEE = AEE.rename(columns = {0: 'Palavra', 1: 'Freq'}, inplace = False)

    print(AEE)

    #pergunta para o usuário responder
    ent = input ("sair => 1 / continuar => 'escrever par de palavras sem sentido'")

    if ent != "1":
        dado = pd.Series(ent)
        rejeitar = rejeitar.append(dado, ignore_index=True)
    else:
        sair = 1

rejeitar.to_excel('Database.xlsx', sheet_name='Regex', index = False)

AEE.to_excel('Resultado.xlsx', sheet_name='Resultado', index = False)

