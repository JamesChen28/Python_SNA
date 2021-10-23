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

data = pd.read_excel('Network_List_v1.0.xlsx', sheet_name='Project_Data')



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
G = nx.MultiGraph() #raph() #MultiGraph()
for i in range(clean.shape[0]):
    G.add_edges_from([(clean['Name_x'][i], clean['Name_y'][i])], Label = clean['Unique-Gkey'][i]) # have to change
    G.add_nodes_from([clean['Name_x'][i]], TYPE = 'daimond')

print(G.edges(data=True))
#### Cal ####
degree = pd.DataFrame(pd.Series(nx.algorithms.centrality.degree_centrality(G))).rename(columns={0: "Degree"})
degree = degree.sort_values(by=['Degree'], ascending=False)
# between = nx.algorithms.centrality.edge_betweenness_centrality(G)
# close = nx.algorithms.centrality.closeness_centrality(G)
# print(degree)
#### Plot ####
pos = nx.spring_layout(G, k=0.8)
edge_labels = {(edge[0], edge[1]): edge[2]['Label'] for edge in G.edges(data=True)}
nx.draw(G,
        pos,
        with_labels=True,
        node_size=1,
        font_size=7,
        width=0.1)
nx.draw_networkx_edge_labels(G,
                             pos,
                             edge_labels = edge_labels,
                             font_size=3,
                             label_pos=0.6)

plt.show()

#### Save ####
if not os.path.exists('Output'):
    os.makedirs('Output')

nx.write_graphml(G,"Output/Network_List_v1.0.graphml")
plt.savefig("Output/Plot_Network_List_v1.0.pdf")
degree.to_excel('Output/Degree_Network_List_v1.0.xlsx', sheet_name = 'degree')

