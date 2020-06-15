#_*_ coding=utf-8 _*_

from graphviz import Digraph
import pandas as pd
import glob

#Associa cada cor na planilha a uma cor do Graphviz
colors = {'Amarelo':'yellow','Azul':'blue','Branco':'white','Cinza':'grey','Marrom':'brown','Ouro':'gold','Preto':'black','Roxo':'purple','Verde':'green','Vermelho':'red','Laranja':'orange'}
font_colors = {'Amarelo':'black','Azul':'white','Branco':'black','Cinza':'black','Marrom':'white','Ouro':'black','Preto':'white','Roxo':'white','Verde':'black','Vermelho':'black','Laranja':'black'}

#Associa os tipos de relação na planilha aos vértices no Graphviz
types = {'Positivo':'blue', 'Negativo':'red', 'Neutro':'black'}
dirs = {'Sim':'both', 'Não':'forward'}

'''
Obtém o data frame do Pandas
'''
def getDataFrame(fileloc):
    #lendo o arquivo EXCEL pelo pandas
    try:
        xls = pd.ExcelFile(fileloc)
    except:
        return None, None

    df_atores = pd.read_excel(xls, 'atores')
    df_relacionamentos = pd.read_excel(xls, 'relacionamentos')

    #pegando os atores e relacionamentos
    df_atores = df_atores.loc[:,['ator', 'cor', 'grupo']].dropna()
    df_relacionamentos = df_relacionamentos.loc[:,['de', 'relacionamento', 'para', 'tipo', 'bilateral']].dropna()
    return df_atores, df_relacionamentos

'''
Faz os nós dos atores
'''
def makeActorNodes(df):
    atores = list(df.iloc[:,0])
    cores = list(df.iloc[:,1])
    grupos = list(df.iloc[:,2])
    group_dict = {}
    for i in range(len(atores)):
        if atores[i] not in grupos:
            g.node(atores[i],shape='circle',fillcolor=colors[cores[i]],style='filled', fontcolor=font_colors[cores[i]], fixedsize='false',width='1')
        if grupos[i] == '-': continue
        if grupos[i] not in group_dict.keys():
            group_dict[grupos[i]] = list()
        group_dict[grupos[i]].append(i)
    return group_dict, atores

'''
Faz os agrupamentos
'''
def makeGroups(group_dict, atores):
    for key in group_dict.keys():
        ids = group_dict[key]
        with g.subgraph(name='cluster_'+key) as c:
            c.attr(label=key)
            c.attr(style='rounded')
            c.attr(rank='min')
            for j in ids:
                c.node(atores[j])

'''
Cria os relacionamentos
'''
def makeRelationships(df, group_dict, atores):
    #ports=['n','ne','e','se','s','sw','w','nw']
    #port_count_h = dict()
    #port_count_t = dict()
    des = list(df.iloc[:,0])
    relacionamentos = list(df.iloc[:,1])
    paras = list(df.iloc[:,2])
    tipos = list(df.iloc[:,3])
    bilaterais = list(df.iloc[:,4])
    for i in range(len(des)):
        #pc_h = port_count_h.get(des[i],0)
        #pc_t = port_count_t.get(des[i],0)

        #Parâmetros básicos
        params = {'dir':dirs[bilaterais[i]], 'color':types[tipos[i]], 'fontcolor':types[tipos[i]], 'fontsize':'10', 'penwidth':'1.0', 'decorate':'false'}

        #De e Para originais
        de = des[i]
        para = paras[i]

        #Ajustando caso o relacionamento envolva clusters
        if des[i] in group_dict.keys():
            #Criando a chave ltail para o relacionamento com o cluster
            params['ltail'] = 'cluster_'+des[i]
            #Pegando um ator qualquer (o primeiro) dentro do cluster
            de = atores[group_dict[des[i]][0]]
        if paras[i] in group_dict.keys():
            #Criando a chave lhead para o relacionamento com o cluster
            params['lhead'] = 'cluster_'+paras[i]
             #Pegando um ator qualquer (o primeiro) dentro do cluster
            para = atores[group_dict[paras[i]][0]]
        g.edge(de, para, label=relacionamentos[i], **params)#, headport=ports[pc_h%8], tailport=ports[pc_t%8])#, decorate='true')
        #port_count_h[des[i]] = pc_h+1
        #port_count_t[des[i]] = pc_t+1

'''
Rodando o código
'''
#Lista arquivos xlsx
xlsx_files=glob.glob('*.xlsx')
#Cria o grafico para cada arquivo encontrado
for fileloc in xlsx_files:
    #ignora arquivos temporários do Excel
    if fileloc.startswith('~$'):
        continue
    
    k = fileloc.rfind(".xlsx")
    g = Digraph(filename=fileloc[:k], engine='dot', format='png')
    g.attr(compound='true')
    g.attr(rankdir='LR')
    g.attr(dpi='600')
    g.attr(ratio = '0.5294')
    g.attr(newrank='true')

    #Constroi os dataframes do excel
    df_atores, df_relacionamentos = getDataFrame(fileloc)
    #Checa se getDataFrame foi bem sucedida
    if df_atores is None and df_relacionamentos is None:
        continue
    #Faz os nós dos atores
    group_dict, atores = makeActorNodes(df_atores)
    #Cria os clusters
    makeGroups(group_dict, atores)
    #Faz os relacionamentos
    makeRelationships(df_relacionamentos, group_dict, atores)
    #Abre o diagrama
    g.view()
