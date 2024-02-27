import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import copy
import win32com.client
import pandas as pd
import json
from datetime import datetime

date = datetime. now(). strftime("%Y_%m_%d-%I%M_%p")

# Create an app object
my_app = _app.connect()

scenario_id = 205
modeller = _m.Modeller(my_app)

# Access the project database
emmebank = modeller.emmebank
# the scenario is accessed from the database by specifying its ID
scenario = emmebank.scenario(scenario_id)
# the network is accessed from the scenario
network = scenario.get_network()


# Functions
# Run Traffic Assignment
def standard_traffic_assignment(numselectedlinks=1):
    # Run standard traffic assignment
    standard_traffic_assignment = modeller.tool("inro.emme.traffic_assignment.sola_traffic_assignment")

    traffic_assignment_spec_file = r'C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\TestProject_emmeLinkAnalysis\Specifications\SOLA_spec.ems'
    with open(traffic_assignment_spec_file) as spec_file:
        traffic_assignment_spec_as_str = spec_file.read()

    traffic_assignment_spec = json.loads(traffic_assignment_spec_as_str)

    traffic_assignment_spec['classes'][0]['path_analyses'][0]['selection_threshold']['upper'] = numselectedlinks
    traffic_assignment_spec['classes'][0]['path_analyses'][0]['selection_threshold']['lower'] = numselectedlinks

    report = standard_traffic_assignment(traffic_assignment_spec)
    spec_file.close()

    return report

# Exiting attribute values reset to 0
def reset_selectlink(select_links,full_linklist):

    network = scenario.get_network()
    # print(full_linklist)
    for l in full_linklist:
        if l in select_links:
            # print(l)
            network.link(l.i_node,l.j_node)['@select'] = 1
        else:
            network.link(l.i_node, l.j_node)['@select'] = 0


    return scenario.publish_network(network)

# Run traffic assignment on non-user select links
def centroid_iter_traffic_assignment(conn_type,network=network,linklist=[],full_link_list=[]):
    """ Connector type can only be access or egress"""
    df_list = []
    if conn_type == "access":
        dir_conn = "OUT"
    else:
        dir_conn = "IN"

    for c in centroid_dict:

        for i in range(0, len(centroid_dict[c][conn_type])):

            for l in linklist:
                network.link(l.i_node, l.j_node)['@select'] = 1

            link = centroid_dict[c][conn_type][i]
            selectlink = "{centroid}_{direction}".format(centroid=c,direction=dir_conn ) + "_" + str(link.i_node) + "-" + str(link.j_node)
            network.link(link.i_node, link.j_node)['@select'] = 1

            scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
            scenario.publish_network(network)

            # Run traffic assignment tool
            standard_traffic_assignment(len(linklist)+1)

            # After running the assignment, the scenario will have updated @select_volumes automatically saved
            # We just need to 'refresh' the network object to fetch these updated volumes
            updated_network = scenario.get_network()

            selectlink_vol_list = []
            link_id_list =[]

            for l in updated_network.links():
                # print(l,";",l["@select_volume"])
                selectlink_vol_list.append(l["@select_volume"])
                link_id_list.append(str(l.i_node) + "-" + str(l.j_node))

            data = list(zip(link_id_list,selectlink_vol_list))

            df = pd.DataFrame(data, columns=["LinkID", selectlink])
            df = df.set_index('LinkID')
            df_list.append(df)
            data.clear()

            reset_selectlink(select_links=linklist,full_linklist=full_link_list)
    print("complete for {conn_type} links".format(conn_type=conn_type))
    return pd.concat(df_list, sort=True)


# Get all centroids and their connectors
## Access is out of zone, egress is into zone
list_centroid = []
user_select_links =[]
linklist = []
centroid_dict = {}

# Get all links in the network
for l in network.links():
    linklist.append(l)

# Get list of user selected links
for l in linklist:
    if network.link(l.i_node,l.j_node)['@select'] == 1:
        user_select_links.append(l)

# Store centroid access/egress links in dictionary
for c in network.centroids():
    list_centroid.append(c)

for c in list_centroid:
    to_centroid_link = []
    from_centroid_link = []
    # Check if links are connected to the centroids
    for l in linklist:
        if l.i_node == c:
            from_centroid_link.append(l)
        if l.j_node == c:
            to_centroid_link.append(l)
    centroid_dict[c] = {'access':copy.copy(from_centroid_link),'egress':copy.copy(to_centroid_link)}

# iterate over each centroid access and egress links separately

dfResultlist = []
for conn_type in ["access", "egress"]:
    network = scenario.get_network()
    dfResultlist.append(centroid_iter_traffic_assignment(conn_type=conn_type,network=network,linklist = user_select_links, full_link_list = linklist))

df_results = pd.concat(dfResultlist ,axis=0, sort=True)
df_results = df_results.groupby(by=["LinkID"],axis=0).sum()
df_results.columns = df_results.columns.str.split('_', expand=True)

df_results_agg = df_results.groupby(level=[0,1],axis=1).sum()
df_results_agg.columns = ['_'.join(col).strip() for col in df_results_agg.columns.values]
df_results_agg.to_csv(r'C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\{date}_Output_agg.csv'.format(date=date))

df_results.columns = ['_'.join(col).strip() for col in df_results.columns.values]
df_results.to_csv(r'C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\{date}_Output_.csv'.format(date=date))