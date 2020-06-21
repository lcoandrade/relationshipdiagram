#_*_ coding=utf-8 _*_

from graphviz import Digraph
import pandas as pd
import glob

#Associa cada cor na planilha a um par (cor de fundo,cor da fonte) do Graphviz
colors = {'Amarelo':('yellow','black'),'Azul':('blue','white'),'Branco':('white','black'),
         'Cinza':('grey','black'),'Marrom':('brown','white'),'Ouro':('gold','black'),
         'Preto':('black','white'),'Roxo':('purple','white'),'Verde':('green','black'),
         'Vermelho':('red','black'),'Laranja':('orange','black')}

#Associa os tipos de relação na planilha aos vértices no Graphviz
rel_types = {'Positivo':'blue', 'Negativo':'red', 'Neutro':'black'}
dirs = {'Sim':'both', 'Não':'forward'}

'''
Obtém o data frame do Pandas
Recebe o nome do arquivo Excel
Retorna as listas de atores e relacionamentos
'''
def get_data(fileloc):
    #lendo o arquivo do Excel pelo pandas
    try:
        df_at_rel = pd.read_excel(fileloc, ['atores','relacionamentos'])
    except:
        return None, None

    #Obtendo os atores e relacionamentos
    df_atores = df_at_rel['atores'].loc[:,['ator', 'cor', 'grupo']].dropna()
    df_relacionamentos = df_at_rel['relacionamentos'].loc[:,['de', 'relacionamento', 'para', 'tipo', 'bilateral']].dropna()
    return df_atores.values.tolist(), df_relacionamentos.values.tolist()

'''
Faz os nós dos atores
Recebe o objeto Digraph e a lista de atores
Retorna o Dicionário com os grupos e membros
'''
def make_actor_nodes(graph,atores):
    grupos = {}
    for ator in atores:
        #expandindo os parâmetros da lista
        ator_nome, ator_cor, ator_grupo = ator
        graph.node(ator_nome,shape='circle',fillcolor=colors[ator_cor][0],style='filled', fontcolor=colors[ator_cor][1], fixedsize='false',width='1')
        if ator_grupo == '-': continue
        if ator_grupo not in grupos:
            grupos[ator_grupo] = list()
        grupos[ator_grupo].append(ator_nome)
    return grupos

'''
Faz os agrupamentos
Recebe o objeto Digraph e Dicionário com os grupos e membros
'''
def makeGroups(graph, grupos):
    for grupo,membros in grupos.items():
        with graph.subgraph(name='cluster_'+grupo) as sub_g:
            sub_g.attr(label=grupo)
            sub_g.attr(style='rounded')
            sub_g.attr(rank='min')
            for membro in membros:
                sub_g.node(membro)

'''
Cria os relacionamentos
Recebe o objeto Digraph, a lista de relacionamentos e o Dicionário com os grupos e membros
'''
def make_relationships(graph, relacionamentos, grupos):
    for relacionamento in relacionamentos:
        #expandindo os parâmetros da lista
        de, legenda, para, tipo, bilateral = relacionamento
        #Parâmetros básicos
        params = {'dir':dirs[bilateral], 'color':rel_types[tipo], 'fontcolor':rel_types[tipo], 'fontsize':'12', 'penwidth':'1.0', 'decorate':'false'}

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
#Lista arquivos xlsx
xlsx_files=glob.glob('*.xlsx')
#Cria o grafico para cada arquivo encontrado
for fileloc in xlsx_files:
    #ignora arquivos temporários do Excel
    if fileloc.startswith('~$'):
        continue
    
    k = fileloc.rfind(".xlsx")
    gattr = {'compound':'true', 'rankdir':'LR', 'dpi':'400', 'ratio':'0.5625', 'newrank':'true'}
    g = Digraph(filename=fileloc[:k], engine='dot', format='png', graph_attr=gattr)

    #Constroi os dataframes do excel
    atores, relacionamentos = get_data(fileloc)
    #Checa se get_data foi bem sucedida
    if atores is None and relacionamentos is None:
        continue
    #Faz os nós dos atores
    grupos = make_actor_nodes(g, atores)
    #Cria os clusters
    makeGroups(g, grupos)
    #Faz os relacionamentos
    make_relationships(g, relacionamentos, grupos)
    #Abre o diagrama
    g.view()
