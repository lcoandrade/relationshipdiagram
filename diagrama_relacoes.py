from graphviz import Digraph
import pandas as pd

#Atores
fileloc = 'atores_relacoes.xlsx'
ator = 'ator'
cor = 'cor'
grupo = 'grupo'
#Relacionamentos
de = 'de'
relacionamento = 'relacionamento'
para = 'para'
tipo = 'tipo'
bilateral = 'bilateral'

#Colors
colors = {'Amarelo':'yellow','Azul':'blue','Cinza':'grey','Marrom':'brown','Ouro':'gold','Preto':'black','Roxo':'purple','Verde':'green','Vermelho':'red','Laranja':'orange'}
font_colors = {'Amarelo':'black','Azul':'white','Cinza':'black','Marrom':'white','Ouro':'black','Preto':'white','Roxo':'white','Verde':'black','Vermelho':'black','Laranja':'black'}

#Relation types
types = {'Positivo':'blue', 'Negativo':'red', 'Neutro':'black'}
dirs = {'Sim':'both', 'Não':'forward'}

'''
Obtém o data frame do Pandas
'''
def getDataFrame(fileloc, ator, cor, grupo, de, relacionamento, para, tipo, bilateral):
    #lendo o arquivo CSV pelo pandas
    xls = pd.ExcelFile(fileloc)
    df_atores = pd.read_excel(xls, 'atores')
    df_relacionamentos = pd.read_excel(xls, 'relacionamentos')

    #pegando os atores e relacionamentos
    df_atores = df_atores.loc[:,[ator, cor, grupo]].dropna()
    df_relacionamentos = df_relacionamentos.loc[:,[de, relacionamento, para, tipo, bilateral]].dropna()
    return df_atores, df_relacionamentos

#Cria o grafico
g = Digraph(filename='diagrama_relacoes', engine='dot', format='png')
g.attr(compound='true')
g.attr(rankdir='BT')
g.attr(dpi='600')
g.attr(ratio = '0.5294')

'''
Faz os nós dos atores
'''
def makeActorNodes(df):
    atores = list(df.iloc[:,0])
    cores = list(df.iloc[:,1])
    grupos = list(df.iloc[:,2])
    group_dict = {}
    for i in range(len(atores)):
        g.node(atores[i],shape='circle',color=colors[cores[i]],style='filled', fontcolor=font_colors[cores[i]], fixedsize='true',width='1')
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
            for j in ids:
                c.node(atores[j])

'''
Cria os relacionamentos
'''
def makeRelationships(df):
    ports=['n','ne','e','se','s','sw','w','nw']
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
        g.edge(des[i], paras[i], dir=dirs[bilaterais[i]], label=relacionamentos[i], color=types[tipos[i]], fontcolor=types[tipos[i]],fontsize = '10', penwidth='1.0')#, headport=ports[pc_h%8], tailport=ports[pc_t%8])#, decorate='true')
        #port_count_h[des[i]] = pc_h+1
        #port_count_t[des[i]] = pc_t+1

#Rodando o código
df_atores, df_relacionamentos = getDataFrame(fileloc, ator, cor, grupo, de, relacionamento, para, tipo, bilateral)
group_dict, atores = makeActorNodes(df_atores)
makeRelationships(df_relacionamentos)
makeGroups(group_dict, atores)
g.view()
