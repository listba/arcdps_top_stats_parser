import os.path
from os import listdir
import json
import gzip

input_directory = 'D:\\GW2_Logs\\Output\\'
files = listdir(input_directory)
sorted_files = sorted(files)

#List of tags from the logs
Tags = {}

for filename in sorted_files:
    # skip files of incorrect filetype
    file_start, file_extension = os.path.splitext(filename)
    #if args.filetype not in file_extension or "top_stats" in file_start:
    if file_extension not in ['.json', '.gz'] or "top_stats" in file_start:
        continue

    print_string = "parsing "+filename
    file_path = "".join((input_directory,"/",filename))

    if file_extension == '.gz':
        with gzip.open(file_path, mode="r") as f:
            json_data = json.loads(f.read().decode('utf-8'))
    else:
        json_datafile = open(file_path, encoding='utf-8')
        json_data = json.load(json_datafile)    
    
    players = json_data['players']
    #Find player with commanderTag then move to next log
    for player in players:
        if player['hasCommanderTag']:
            print_string += "   -   "+ player['name']
            if player['name'] in Tags:
                Tags[player['name']] += 1
            else:
                Tags[player['name']] = 1
            continue
    
    print(print_string)

Tag_in_Logs = ''
i=0
for Tag in Tags:
    if i >0:
        Tag_in_Logs = Tag_in_Logs+ ', '+Tag+"("+str(Tags[Tag])+" Fights)"
    else:
        Tag_in_Logs = Tag_in_Logs+ ' '+Tag+"("+str(Tags[Tag])+" Fights)"
    i+=1
print('WVW Log Review with'+Tag_in_Logs)
#json_datafile.close()
