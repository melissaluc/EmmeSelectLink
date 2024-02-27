import inro.emme.desktop.app as _app
import inro.modeller as _m
import os
import win32com.client
from pathlib import Path



# Emme application connect to current Emme instance
my_app = _app.connect()

scenario_id = 205
modeller = _m.Modeller(my_app)

# Access the project database
emmebank = modeller.emmebank
# the scenario is accessed from the database by specifying its ID
scenario = emmebank.scenario(scenario_id)
# the network is accessed from the scenario
network = scenario.get_network()


# #  Option 1: Read links in Excel into list

# Current working directory
wd = os.getcwd()
fp = Path(wd)

print(wd)

for f in list(fp.glob('*')):
    if os.path.basename(f) == 'Link Lists.xlsx':
        fp_ll = f

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
wb_l = xl.Workbooks.Open(fp_ll)
ws_l = wb_l.Worksheets(1)

link_list = []

for r in range(1,10):
    if ws_l.Cells(r,1).Value!= None:
        if "-" in ws_l.Cells(r,1).Value:
            link_list.append(ws_l.Cells(r,1).Value)

# Create extra attribute @select and @select_volume
extra_attribute_list = []
for extra_attribute in scenario.extra_attributes():
    extra_attribute_list.append(extra_attribute.name)

print(extra_attribute_list)

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
NAMESPACE = "inro.emme.data.extra_attribute.init_extra_attribute"
init_extra = _m.Modeller().tool(NAMESPACE)

for att in ['@select','@select_volume']:
    a = _m.Modeller().scenario.extra_attribute(att)
    init_extra(a, 0)


# # Get new set attribute values
# att_values_select = scenario.get_attribute_values('LINK', ['@select'])
# scenario.set_attribute_values('LINK', ['@select'], att_values_select)
#
# att_values_selectvol = scenario.get_attribute_values('LINK', ['@select_volume'])
# scenario.set_attribute_values('LINK', ['@select_volume'], att_values_selectvol)


# Get link id in Emme, if link is in excel list then set to 1
for l in network.links():
    if l.id in link_list:
        print(f"{l.id} is in link list")
        network.link(l.i_node,l.j_node)['@select'] = 1
        print('The value is set to: {0}'.format(network.link(l.i_node,l.j_node)['@select']))


# # Set all values assigned to 1
# att_values_select_2 = scenario.get_attribute_values('LINK', ['@select'])
# scenario.set_attribute_values('LINK', ['@select'], att_values_select_2)

# Publish the edits to the network
scenario.publish_network(network)

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


select_vol = scenario.get_attribute_values('LINK', ['@select_volume'])
scenario.set_attribute_values('LINK',['@select_volume'],select_vol)

for i in scenario.get_attribute_values('LINK', ['@select_volume']):
    print(i)

print(network.link("121","10952")['@select_volume'])
# Publish the edits to the network
scenario.publish_network(network)

print(network.link("12132","11127")['@select'])
print(network.link("12132","11127")['@select_volume'])
# Write volau results to excel

# # Excel application
# excelobj = win32com.client.Dispatch('Excel.Application')
# excelobj.Visible = True
# excelwrkbk = excelobj.Workbooks.Add()
# sh = excelobj.ActiveSheet
#
# sh.Cells(1, 1).Value = "Link Id"
# sh.Cells(1, 2).Value = "@select_volume"
#
# r = 2
# for l in network.links():
#     # print(l.id)
#     sh.Cells(r, 1).Value = l.id
#     # print(network.link(l.i_node,l.j_node)[selectedlinkvolumeatt])
#     sh.Cells(r, 2).Value = network.link(l.i_node,l.j_node)[selectedlinkvolumeatt]
#     r += 1
