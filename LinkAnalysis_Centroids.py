import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import win32com.client
from pathlib import Path
import copy

# Functions
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
                "lower": 1,
                "upper": 1
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
    return report, print(type(report))


# Emme application connect to current Emme instance
emme_app = _app.connect()

scenario_id = 205
modeller = _m.Modeller(emme_app)

# Access the project database
emmebank = modeller.emmebank
# the scenario is accessed from the database by specifying its ID
scenario = emmebank.scenario(scenario_id)
# the network is accessed from the scenario
network = scenario.get_network()

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
for link in network.links():
    network.link(link.i_node, link.j_node)['@select'] = 0

att_values_select = scenario.get_attribute_values('LINK', ['@select'])
scenario.set_attribute_values('LINK', ['@select'], att_values_select)

# # Set link
# # network.link('11145','11148')['@select'] = 0
# network.link('10030', '10063')['@select'] = 0
#

# # Get link id
# for link in network.links():
#     if network.link(link.i_node, link.j_node)['@select'] == 1:
#         print(f'Select link analysis for link: {link.id}')

# Get all centroids and their connectors
## Access is out of zone, egress is into zone
list_centroid = []
for c in network.centroids():
    list_centroid.append(c)

centroid_dict = {}

# Get all links in the network
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


for c in centroid_dict:

    # print(c,"----",centroid_dict[c]['access'][2])
    for i in range(0, len(centroid_dict[c]['access'])):
        print(c,": ",centroid_dict[c]['access'][i])
        aclink = centroid_dict[c]['access'][i]
        network.link(aclink.i_node, aclink.j_node)['@select'] = 1

        scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
        scenario.publish_network(network, resolve_attributes=True)
        # Run traffic assignment tool
        # Run standard traffic assignment

        for link in network.links():
            if network.link(link.i_node, link.j_node)['@select'] == 1:
                    print(f'Select link analysis for link: {link.id}')
                    standard_traffic_assignment()

        scenario.set_attribute_values('LINK', ['@select_volume'], scenario.get_attribute_values('LINK', ['@select_volume']))

        # # Reset connector link to 0
        network.link(aclink.i_node, aclink.j_node)['@select'] = 0
        scenario.set_attribute_values('LINK', ['@select'], scenario.get_attribute_values('LINK', ['@select']))
        scenario.publish_network(network, resolve_attributes=True)

        # Store results
        print(scenario.get_attribute_values('LINK', ['@select_volume']))


    # for i in range(0,len(centroid_dict[c]['egress'])):
    #     # print(c,": ",centroid_dict[c]['access'][i])
    #     eclink = centroid_dict[c]['egress'][i]
    #     network.link(eclink.i_node,eclink.j_node)['@select'] = 1
    #     print(c,"-egress-",network.link(eclink.i_node,eclink.j_node)['@select'])
    #     standard_traffic_assignment()
    #     sh.Cells(1, col).Value = f"{c} IN"
    #     sh.Cells(row, col).Value = network.link(aclink.i_node,aclink.j_node)['@select_volume']
    #     col += 1
    #     # Exiting attribute values reset to 0
    #     for link in network.links():
    #         network.link(link.i_node,link.j_node)['@select'] = 0

    #     att_values_select = scenario.get_attribute_values('LINK', ['@select'])
    #     scenario.set_attribute_values('LINK', ['@select'], att_values_select)
