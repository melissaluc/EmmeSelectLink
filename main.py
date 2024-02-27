import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import copy

def unique(list1):
    # intilize a null list
    unique_list = []

    # traverse for all elements
    for x in list1:
        # check if exists in unique_list or not
        if x not in unique_list:
            unique_list.append(x)
    return unique_list

my_app = _app.connect()

scenario_id = 205
modeller = _m.Modeller(my_app)

# we access the project database
emmebank = modeller.emmebank
# the scenario is accessed from the database by specifying its ID
scenario = emmebank.scenario(scenario_id)
# the network is accessed from the scenario
network = scenario.get_network()

# list_centroid = []
# for c in network.centroids():
#     list_centroid.append(c)
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


# Create extra attribute @select
extra_attribute_list = []
for extra_attribute in scenario.extra_attributes():
    extra_attribute_list.append(extra_attribute.name)

# Check if @select exists and set to 0
if "@select" not in extra_attribute_list:
    scenario.create_extra_attribute('LINK', '@select', 0)

for l in network.links():
    network.link(l.i_node,l.j_node)['@select'] = 0

selectsetatt_values = scenario.get_attribute_values('LINK', ['@select'])
scenario.set_attribute_values('LINK', ['@select'], selectsetatt_values)


# Set centroid links to 1, read link nodes in excel inode, jnode
network.link("10032","12051")['@select']=1


# set all link values
values = scenario.get_attribute_values('LINK', ['@select'])
scenario.set_attribute_values('LINK', ['@select'], values)
# Publish the edits to the network
scenario.publish_network(network)

# Run standard traffic assignment
standard_traffic_assignment = modeller.tool("inro.emme.traffic_assignment.standard_traffic_assignment")

vmode = 'c'
mfdemand = 'mf1'
selectedlinkvolumeatt = '@select_volume'
lthresh,uthresh = 1,1

spec = '''
{
    "type": "STANDARD_TRAFFIC_ASSIGNMENT",
    "classes": [
        {
            "mode": "%s",
            "demand": "%s",
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
                    "selected_link_volumes": "%s",
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
            "lower": %s,
            "upper": %s
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
}'''% (vmode, mfdemand, selectedlinkvolumeatt,lthresh,uthresh)

print(network.link("10032","12051")['@select'])


report = standard_traffic_assignment(spec)
print(network.link("10032","12051")['@select_volume'])
# Publish the edits to the network
scenario.publish_network(network)