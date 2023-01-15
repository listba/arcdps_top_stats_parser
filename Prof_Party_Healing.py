from dataclasses import dataclass,field
import os.path
from os import listdir
from enum import Enum
from xlutils.copy import copy
import json
import gzip

from collections import OrderedDict

input_directory = 'D:\\GW2_Logs\\Output\\'
files = listdir(input_directory)
sorted_files = sorted(files)

def my_value(number):
    return ("{:,}".format(number))

OutgoingHealing = {}

for filename in sorted_files:
    # skip files of incorrect filetype
    file_start, file_extension = os.path.splitext(filename)
    #if args.filetype not in file_extension or "top_stats" in file_start:
    if file_extension not in ['.json', '.gz'] or "top_stats" in file_start:
        continue

    print_string = "parsing "+filename
    print(print_string)
    file_path = "".join((input_directory,"/",filename))

    if file_extension == '.gz':
        with gzip.open(file_path, mode="r") as f:
            json_data = json.loads(f.read().decode('utf-8'))
    else:
        json_datafile = open(file_path, encoding='utf-8')
        json_data = json.load(json_datafile)    
    
    players = json_data['players']
    

    for player in players:
        if 'extHealingStats' in player:
            healerName = player['name']
            healerProf = player['profession']
            healerGroup = player['group']
            if healerProf not in OutgoingHealing:
                OutgoingHealing[healerProf] = {}
                OutgoingHealing[healerProf]['inPartyHeals'] = 0
                OutgoingHealing[healerProf]['outPartyHeals'] = 0
                OutgoingHealing[healerProf]['inPartyBarrier'] = 0
                OutgoingHealing[healerProf]['outPartyBarrier'] = 0
                
            
            for index, target in enumerate(player['extHealingStats']['outgoingHealingAllies']):
                    targetHealing = target[0]['healing']
                    targetName = players[index]['name']
                    targetProf = players[index]['profession']
                    targetGroup = players[index]['group']
                    if targetGroup == healerGroup:
                        OutgoingHealing[healerProf]['inPartyHeals'] += targetHealing
                    else:
                        OutgoingHealing[healerProf]['outPartyHeals'] += targetHealing
                    
print("|Healer Prof | Total Healing| In Party Heals| In Party %|Out Party Heals| Out Party %|h")
output_string = ""
for prof in OutgoingHealing:
    Total_Heals = (OutgoingHealing[prof]['inPartyHeals'] + OutgoingHealing[prof]['outPartyHeals'])
    Total_Barrier = (OutgoingHealing[healerProf]['inPartyBarrier'] + OutgoingHealing[healerProf]['outPartyBarrier'])
    In_Heals = 0
    Out_Heals = 0
    H_In_Percentage = 0.00
    H_Out_Percentage = 0.00
    if Total_Heals >0:
        In_Heals = OutgoingHealing[prof]['inPartyHeals']
        Out_Heals = OutgoingHealing[prof]['outPartyHeals']
        if In_Heals > 0:
            H_In_Percentage = round(((In_Heals/Total_Heals)*100),2)
        else:
            H_In_Percentage = 0.00
        if Out_Heals > 0:
            H_Out_Percentage = round(((Out_Heals/Total_Heals)*100),2)
        else:
            H_Out_Percentage = 0.00
    if Total_Heals >0:
        print("|"+prof+" | "+my_value(Total_Heals)+"| "+my_value(In_Heals)+"| "+str(H_In_Percentage)+"%| "+my_value(Out_Heals)+"| "+str(H_Out_Percentage)+"%|")


                            
#json_datafile.close()
