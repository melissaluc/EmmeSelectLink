import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import copy
import win32com.client
import pandas as pd
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
def standard_traffic_assignment():
    # Run standard traffic assignment
    standard_traffic_assignment = modeller.tool("inro.emme.traffic_assignment.standard_traffic_assignment")


    spec = '''{
        "type": "STANDARD_TRAFFIC_ASSIGNMENT",
        "classes": [
            {
                "mode": "c",
                "demand": "mf1",
                "generalized_cost": null,
                "results": {
                    "link_volumes": null,
                    "turn_volumes": null,
                    "od_travel_times": {
                        "shortest_paths": null
                    }
                },
                "analysis": {
                    "analyzed_demand": null,
                    "results": {
                        "od_values": null,
                        "selected_link_volumes": "@select_volume",
                        "selected_turn_volumes": null
                    }
                }
            }
        ],
        "performance_settings": {
            "number_of_processors": "max"
        },
        "background_traffic": null,
        "path_analysis": {
            "link_component": "@select",
            "turn_component": null,
            "operator": "+",
            "selection_threshold": {
                "lower": 2,
                "upper": 10
            },
            "path_to_od_composition": {
                "considered_paths": "ALL",
                "multiply_path_proportions_by": {
                    "analyzed_demand": false,
                    "path_value": true
                }
            }
        },
        "cutoff_analysis": null,
        "traversal_analysis": null,
        "stopping_criteria": {
            "max_iterations": 100,
            "relative_gap": 0,
            "best_relative_gap": 0.1,
            "normalized_gap": 0.05
        }
    }'''

    report = standard_traffic_assignment(spec)
    return report

# Exiting attribute values reset to 0
def reset_selectlink(network,scenario):

    for link in network.links():
        network.link(link.i_node,link.j_node)['@select'] = 0
    

    att_values_select = scenario.get_attribute_values('LINK', ['@select'])
    scenario.set_attribute_values('LINK', ['@select'], att_values_select)
    return scenario.publish_network(network)

# Create extra attribute @select and @select_volume
extra_attribute_list = []
for extra_attribute in scenario.extra_attributes():
    extra_attribute_list.append(extra_attribute.name)


    # Check if @select exists and set to 0
if "@select" not in extra_attribute_list:
    scenario.create_extra_attribute('LINK', '@select', 0)
    print("Created extra link attribute @select")
else:
    print("@select already exists")

if "@select_volume" not in extra_attribute_list:
    scenario.create_extra_attribute('LINK', '@select_volume', 0)
    print("Created extra link attribute @select_volume")
else:
    print("@select_volume already exists")

# Exiting attribute values reset to 0
reset_selectlink(network,scenario)

network.link('10032','12051')['@select'] = 1
scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
scenario.publish_network(network)

# Get all centroids and their connectors
## Access is out of zone, egress is into zone
list_centroid = []
for c in network.centroids():
    list_centroid.append(c)

centroid_dict = {}

## Get all links in the network
linklist = []
for l in network.links():
    linklist.append(l)

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

# List of Links to gather data from

# Run Traffic Assignment for each centroid separating IN and OUT trips

df_list=[]

data = {}
for c in centroid_dict:
# Set all access links to 1
# Reset Access links to 0
# Set all egress links to 1
# Reset Access links to 0
    network = scenario.get_network()
## Access links - from centroid therefore OUT
    # Set connector links leaving zone to 1
    
    for i in range(0,len(centroid_dict[c]['access'])):
        aclink = centroid_dict[c]['access'][i]
        selectlink ="{centroid}_OUT"
        network.link(aclink.i_node,aclink.j_node)['@select'] = 1
        
        scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
        scenario.publish_network(network)

    # Run traffic assignment tool
    standard_traffic_assignment()

    # After running the assignment, the scenario will have updated @select_volumes automatically saved
    # We just need to 'refresh' the network object to fetch these updated volumes
    network= scenario.get_network()

    # print(scenario.get_attribute_values('LINK', ['@select_volume'])[1],"/n--------------------------------------------------------------------")
    f = open(r"C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\STA_Result_OUT_[{centroid}]_{selectlink}.txt".format(centroid = c, selectlink = aclink), "w")
    
        
    
    for link in network.links():
        f.write("{0};{1}\n".format(link.id,link["@select_volume"]))
        link_id = str(link.i_node)+"-"+str(link.j_node)
        if link_id in data.keys():
                data[link_id].append(link["@select_volume"])
        else:
            data[link_id] = link["@select_volume"]
                
    selectlink ="{centroid}_OUT".format(centroid=c)
    df_OUT = pd.DataFrame(data.items(),columns=["LinkID",selectlink])
    df_OUT = df_OUT.set_index('LinkID')
    df_list.append(df_OUT)
    data.clear()
    reset_selectlink(network,scenario)

    if network.link(aclink.i_node,aclink.j_node)['@select'] == 1:
        print("Did not reset links to 0")
    else:
        print("all links for @select is set to 0")
    network.link('10032','12051')['@select'] = 1
    scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
    scenario.publish_network(network)


# ## Egress links - to centroid therefore IN
    for i in range(0,len(centroid_dict[c]['egress'])):
            # print(c,": ",centroid_dict[c]['egress'][i])
            eglink = centroid_dict[c]['egress'][i]
            network.link(eglink.i_node,eglink.j_node)['@select'] = 1
            
            scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
            scenario.publish_network(network)


    # Run traffic assignment tool
    standard_traffic_assignment()

    # After running the assignment, the scenario will have updated @select_volumes automatically saved
    # We just need to 'refresh' the network object to fetch these updated volumes
    network = scenario.get_network()

    # print(scenario.get_attribute_values('LINK', ['@select_volume'])[1],"/n--------------------------------------------------------------------")
    f = open(r"C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\STA_Result_IN_[{centroid}]_{selectlink}.txt".format(centroid = c, selectlink = eglink), "w")
          
        
    for link in network.links():
        f.write("{0};{1}\n".format(link.id,link["@select_volume"]))
        link_id = str(link.i_node)+"-"+str(link.j_node)
        if link_id in data.keys():
            data[link_id].append(link["@select_volume"])
        else:
            data[link_id] = link["@select_volume"]
                    
    selectlink ="{centroid}_IN".format(centroid = c)    
    df_IN = pd.DataFrame(data.items(),columns=["LinkID",selectlink])
    df_IN = df_IN.set_index('LinkID')
    df_list.append(df_IN)
    data.clear()
    reset_selectlink(network,scenario)

    if network.link(eglink.i_node,eglink.j_node)['@select'] == 1:
        print("Did not reset links to 0 for ",network.link(eglink.i_node,eglink.j_node))
    else:
        print("all links for @select is set to 0")

    network.link('10032','12051')['@select'] = 1
    scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
    scenario.publish_network(network)


df_results = pd.concat(df_list,axis=1,sort=True)
df_results = pd.concat(df_list,axis=1,sort=True)
df_results.columns = df_results.columns.str.split('_', expand=True)
df_results.to_csv(r'C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\Output.csv')


df_results.to_csv(r'C:\Users\melissa.luc\Desktop\Testing Emme Link Selection Script\Output\Output_Aggregated.csv')

