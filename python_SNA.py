# -*- coding: utf-8 -*-
"""
Created on Sun Oct 24 01:38:10 2021

@author: James
"""

# import package
import os
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt


# import data
path = r'D:\James\Desktop\Coding\Python\network'
os.chdir(path)

filename = 'Network_List_v1.0'
data = pd.read_excel(filename + '.xlsx', sheet_name = 'Project_Data')
data = data.dropna()


#### for Client Project ####
# clean = data.merge(data, on=['Client', 'Project', 'Client Project'], how = 'outer')
# clean = pd.DataFrame(clean[clean.Name_x != clean.Name_y]).reset_index()
# clean = pd.DataFrame(clean[pd.isna(clean.Name_x) == False]).reset_index()

#### for Unique-Gkey ####
clean = data.merge(data, on=['Client', 'Project', 'Unique-Gkey'], how = 'outer')
clean = pd.DataFrame(clean[clean.Name_x != clean.Name_y])
clean = pd.DataFrame(clean[pd.isna(clean.Name_x) == False]).reset_index()
# print(clean)
# clean.to_excel('./clean_Dev.xlsx', sheet_name = 'clean')

#### graph ####
G = nx.DiGraph() #Graph() #MultiGraph()
for i in range(clean.shape[0]):
    G.add_edges_from([(clean['Name_x'][i], clean['Name_y'][i])], Label = clean['Unique-Gkey'][i]) # have to change
    G.add_nodes_from([clean['Name_x'][i]], TYPE = 'daimond')

# print(G.edges(data=True))

#### Cal ####
degree = pd.DataFrame(pd.Series(nx.algorithms.centrality.degree_centrality(G))).rename(columns = {0: "Degree"})
degree = degree.sort_values(by = ['Degree'], ascending = False)
# between = nx.algorithms.centrality.edge_betweenness_centrality(G)
# between = nx.algorithms.centrality.betweenness_centrality(G)
between = pd.DataFrame(pd.Series(nx.algorithms.centrality.betweenness_centrality(G))).rename(columns = {0: "Betweenness"})
between = between.sort_values(by = ['Betweenness'], ascending = False)
close = pd.DataFrame(pd.Series(nx.algorithms.centrality.closeness_centrality(G))).rename(columns = {0: "Closeness"})
close = close.sort_values(by = ['Closeness'], ascending = False)

# print(degree)

center = degree.merge(between, how = 'left', left_index = True, right_index = True).merge(close, how = 'left', left_index = True, right_index = True)


#### Plot ####
pos = nx.spring_layout(G, k = 0.8)
edge_labels = {(edge[0], edge[1]): edge[2]['Label'] for edge in G.edges(data = True)}
nx.draw(G, pos, with_labels = True, node_size = 1, font_size = 10, width = 0.1)
nx.draw_networkx_edge_labels(G, pos, edge_labels = edge_labels, font_size = 5, label_pos = 0.6)

plt.show()

#### Save ####
if not os.path.exists('Output'):
    os.makedirs('Output')

# nx.write_graphml(G,'Output/' + filename+ '.graphml')
plt.savefig('Output/Plot_' + filename+ '.pdf')
center.to_excel('Output/Centrality_' + filename+ '.xlsx', sheet_name = 'centrality')

