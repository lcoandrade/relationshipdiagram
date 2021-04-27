#_*_ coding=utf-8 _*_

from graphviz import Digraph
import pandas as pd
import numpy as np
import glob

#Associa cada cor na planilha a um par (cor de fundo,cor da fonte) do Graphviz
colors = {'Amarelo':('yellow','black'),'Azul':('blue','white'),'Branco':('white','black'),
         'Cinza':('grey','black'),'Marrom':('brown','white'),'Ouro':('gold3','black'),
         'Preto':('black','white'),'Roxo':('purple','white'),'Verde':('green','black'),
         'Vermelho':('red','black'),'Laranja':('orangered','black')}
colors_sub = { 'Amarelo':'#ffff005f','Azul':'cyan','Branco':'white','Cinza':'gray95',
             'Marrom':'#a52a2a5f','Ouro':'gold','Preto':'black','Roxo':'mediumpurple1',
             'Verde':'palegreen','Vermelho':'tomato','Laranja':'orange' }

#Associa os tipos de relação na planilha aos vértices no Graphviz
rel_types = {'Positivo':'blue', 'Negativo':'red', 'Neutro':'black'}
dirs = {'Sim':'both', 'Não':'forward'}

'''
Obtém o data frame do Pandas
Recebe o nome do arquivo Excel
Retorna as listas de atores e relacionamentos
'''
def get_data(fileloc, ext):
    #lendo o arquivo usando Pandas
    ext_engine={'.xls':'xlrd','.xlsx':'openpyxl','.ods':'odf'}
    try:
        df_at_rel = pd.read_excel(fileloc, ['atores','relacionamentos'], engine=ext_engine.get(ext,None))
    except:
        return None, None

    #Obtendo os atores e relacionamentos
    df_atores = df_at_rel['atores'].loc[:,['ator', 'cor', 'grupo']].dropna(subset=['ator']).replace({np.nan: None})
    df_relacionamentos = df_at_rel['relacionamentos'].loc[:,['de', 'relacionamento', 'para', 'tipo', 'bilateral']].dropna()
    return df_atores.to_numpy().tolist(), df_relacionamentos.to_numpy().tolist()

'''
Faz os nós dos atores
Recebe o objeto Digraph e a lista de atores
Retorna o Dicionário com os grupos e membros
'''
def make_actor_nodes(graph,atores):
    lst_grupos = set()
    for ator in atores:
        ator_nome, ator_cor, ator_grupo = ator
        if ator_grupo is not None: lst_grupos.add(ator_grupo)
    grupos = {}
    grupos_cores = {}
    for ator in atores:        
        #expandindo os parâmetros da lista
        ator_nome, ator_cor, ator_grupo = ator
        if ator_cor is None: ator_cor = "Branco"
        if ator_nome in lst_grupos:
            grupos_cores[ator_nome] = ator_cor
        else:
            graph.node(ator_nome,shape='circle',fillcolor=colors[ator_cor][0],style='filled', fontcolor=colors[ator_cor][1], fixedsize='false',width='1')
        if ator_grupo is None: continue
        if ator_grupo not in grupos:
            grupos[ator_grupo] = list()
        grupos[ator_grupo].append(ator_nome)
    return grupos, grupos_cores

'''
Faz os agrupamentos
Recebe o objeto Digraph e Dicionário com os grupos e membros
'''
def makeGroups(graph, grupos, grupos_cores):
    sub_graphs = {}
    for grupo,membros in grupos.items():
        sub_g = Digraph(name='cluster_'+grupo)
        sub_g.attr(label=grupo)
        sub_g.attr(style='rounded')
        sub_g.attr(bgcolor=colors_sub[grupos_cores.get(grupo,'Branco')])
        sub_g.attr(rank='source')
        for membro in membros:
            if membro not in grupos.keys():
                sub_g.node(membro)
        sub_graphs[grupo]=sub_g

    for grupo,membros in grupos.items():
        for membro in membros:
            if membro in grupos.keys():
                sub_graphs[grupo].subgraph(sub_graphs[membro])
                sub_graphs.pop(membro)
                
    for s in sub_graphs.values():
        graph.subgraph(s)



'''
Cria os relacionamentos
Recebe o objeto Digraph, a lista de relacionamentos e o Dicionário com os grupos e membros
'''
def make_relationships(graph, relacionamentos, grupos):
    for relacionamento in relacionamentos:
        #expandindo os parâmetros da lista
        de, legenda, para, tipo, bilateral = relacionamento
        #Parâmetros básicos
        params = {'dir':dirs[bilateral], 'color':rel_types[tipo], 'fontcolor':rel_types[tipo], 'penwidth':'1.0', 'decorate':'false', 'minlen':'2'}

        #Ajustando caso o relacionamento envolva clusters
        if de in grupos.keys():
            #Criando a chave ltail para o relacionamento com o cluster
            params['ltail'] = 'cluster_'+de
            #Pegando um ator qualquer (o primeiro) dentro do cluster
            deN = grupos[de][0]
        else:
            deN = de
        if para in grupos.keys():
            #Criando a chave lhead para o relacionamento com o cluster
            params['lhead'] = 'cluster_'+para
             #Pegando um ator qualquer (o primeiro) dentro do cluster
            paraN = grupos[para][0]
        else:
            paraN = para
        graph.edge(deN, paraN, label=legenda, **params)

'''
Rodando o código
'''
#Lista arquivos xls, xlsx e ods
files = {f:glob.glob('*'+f) for f in ['.ods','.xls','.xlsx']}

#Cria o gráfico para cada arquivo encontrado
for ext,lista_f in files.items():
    for fileloc in lista_f:
        #ignora arquivos temporários do Excel
        if fileloc.startswith('~$'):
            continue
        
        k = fileloc.rfind(ext)
        gattr = dict()
        gattr['compound'] = 'true'
        gattr['rankdir'] = 'LR'
        gattr['dpi'] = '300'
        gattr['ratio'] = '0.5625'
        gattr['newrank'] = 'true'
        gattr['overlap'] = 'false'
        gattr['splines'] = 'true'
        gattr['fontsize'] = '20'

        g = Digraph(filename=fileloc[:k], engine='dot', format='png', graph_attr=gattr)

        #Constroi os dataframes do excel
        atores, relacionamentos = get_data(fileloc, ext)
        #Checa se get_data foi bem sucedida
        if atores is None and relacionamentos is None:
            continue
        #Faz os nós dos atores
        grupos, grupos_cores = make_actor_nodes(g, atores)
        #Cria os clusters
        makeGroups(g, grupos, grupos_cores)
        #Faz os relacionamentos
        make_relationships(g, relacionamentos, grupos)
        #Abre o diagrama
        g.view()
