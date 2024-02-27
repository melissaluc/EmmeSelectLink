import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import copy
import win32com.client
import requests
import pprint
import re


excelobj = win32com.client.Dispatch('Excel.Application')
excelobj.Visible = False
excelwrkbk = excelobj.Workbooks.Add()
sh = excelobj.ActiveSheet


my_app = _app.connect()

scenario_id = 6
#scenario_id = input('Scenario ID:')
modeller = _m.Modeller(my_app)

# we access the project database
emmebank = modeller.emmebank
# the scenario is accessed from the database by specifying its ID
scenario = emmebank.scenario(scenario_id)

# the network is accessed from the scenario
network = scenario.get_network()




# f = open(r'C:\Users\melissa.luc\Desktop\GIS Stuff\nodes_wLabels.txt','w')
# for n in network.nodes():
#     if n.label == "QAZ":
#         f.write(f"{n.id};\n")
#
# # get all nodes
# f = open(r'C:\Users\melissa.luc\Desktop\GIS Stuff\nodesHH_wLabels.txt','w')
# for n in network.nodes():
#     f.write(f"{n.id};\n")
#
# f = open(r'C:\Users\melissa.luc\Desktop\GIS Stuff\linksGTA_wLabels.txt','w')
# for l in network.links():
#     if network.link(l.i_node,l.j_node)['type']==123:
#         f.write(f"{l.id},\n")
#


# #for l in network.links():
# #    if l.i_node in list_centroid:
# #        from_centroid_link.append(l)
# #    if l.j_node in list_centroid:
# #        to_centroid_link.append(l)
#
# centroid_dict = {}
#
# # Get all links in the network
# linklist = []
# for l in network.links():
#     linklist.append(l)
#
# for c in list_centroid:
#
#     to_centroid_link = []
#     from_centroid_link = []
#     # Check if links are connected to the centroids
#     for l in linklist:
#         if l.i_node == c:
#             from_centroid_link.append(l)
#         if l.j_node == c:
#             to_centroid_link.append(l)
#     centroid_dict[c] = {'access':copy.copy(from_centroid_link),'egress':copy.copy(to_centroid_link)}

# node_list = ["33238","33381","33382","41366","41376","41788","42121","43111","43679","43695","43696","43770","43771","43772","50043","51439","51440","51445","51447"]
# count = 0
# for l in network.nodes():
#     if l.id in node_list:
#         network.node(l.id)['label'] = 'QAZ'
#         count +=1


    #
    # if network.node(l.id)['label'] == 'QAZ':
    #     count +=1
    #     print(count)
    #     if network.node(l.id) in node_list:
    #         network.node(l.id)['label'] = 'QAZ'
    #     else:
    #         network.node(l.id)['label'] = "0"



    #if network.node(l.id)['label'] == "QAZ":
    #    network.node(l.id)['label'] = 1


link_ilist = ["4354-42038", "4354-41376","4354-41375","42038-4354","41376-4354","41375-4354"]
#
# linklist_str = []

# for i in link_ilist:
#     linklist_str.append(str(i))
#     print(i)
#
#
count = 0
for l in network.links():
    if l.id in link_ilist:
        network.link(l.i_node,l.j_node)['type'] = 125
        count +=1

print(count)
#
selectsetatt_values = scenario.get_attribute_values('LINK', ['type'])
scenario.set_attribute_values('LINK', ['type'], selectsetatt_values)
scenario.publish_network(network)

#
#
#
# #print(network.link("1001","10000"))
# selectsetatt_values = scenario.get_attribute_values('NODE', ['label'])
# scenario.set_attribute_values('NODE', ['label'], selectsetatt_values)
# scenario.publish_network(network)


#print("{0} and {1} and count:{2}".format(network.node("51602")['label'],network.node("43111")['label'],count))