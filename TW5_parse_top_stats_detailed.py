#!/usr/bin/env python3

#    parse_top_stats_detailed.py outputs detailed top stats in arcdps logs as parsed by Elite Insights.
#    Copyright (C) 2021 Freya Fleckenstein
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.


import argparse
import datetime
import os.path
from os import listdir
import sys
import xml.etree.ElementTree as ET
from enum import Enum
import importlib
import xlwt

from TW5_parse_top_stats_tools import *

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='This reads a set of arcdps reports in xml format and generates top stats.')
    parser.add_argument('input_directory', help='Directory containing .xml or .json files from arcdps reports')
    parser.add_argument('-o', '--output', dest="output_filename", help="Text file to write the computed top stats")
    #parser.add_argument('-f', '--input_filetype', dest="filetype", help="filetype of input files. Currently supports json and xml, defaults to json.", default="json")
    parser.add_argument('-x', '--xls_output', dest="xls_output_filename", help="xls file to write the computed top stats")    
    parser.add_argument('-j', '--json_output', dest="json_output_filename", help="json file to write the computed top stats to")    
    parser.add_argument('-l', '--log_file', dest="log_file", help="Logging file with all the output")
    parser.add_argument('-c', '--config_file', dest="config_file", help="Config file with all the settings", default="TW5_parser_config_detailed")
    parser.add_argument('-a', '--anonymized', dest="anonymize", help="Create an anonymized version of the top stats. All account and character names will be replaced.", default=False, action='store_true')
    args = parser.parse_args()

    if not os.path.isdir(args.input_directory):
        print("Directory ",args.input_directory," is not a directory or does not exist!")
        sys.exit()
    if args.output_filename is None:
        args.output_filename = args.input_directory+"/TW5_top_stats_detailed.tid"
    if args.xls_output_filename is None:
        args.xls_output_filename = args.input_directory+"/TW5_top_stats_detailed.xls"
    if args.json_output_filename is None:
        args.json_output_filename = args.input_directory+"/TW5_top_stats_detailed.json"                
    if args.log_file is None:
        args.log_file = args.input_directory+"/log_detailed.txt"

    output = open(args.output_filename, "w",encoding="utf-8")
    log = open(args.log_file, "w")

    parser_config = importlib.import_module("parser_configs."+args.config_file , package=None) 
    
    config = fill_config(parser_config)

    print_string = "Using input directory "+args.input_directory+", writing output to "+args.output_filename+" and log to "+args.log_file
    print(print_string)
    print_string = "Considering fights with at least "+str(config.min_allied_players)+" allied players and at least "+str(config.min_enemy_players)+" enemies that took longer than "+str(config.min_fight_duration)+" s."
    myprint(log, print_string)

    players, fights, found_healing, found_barrier, squad_comp = collect_stat_data(args, config, log, args.anonymize)    

    # create xls file if it doesn't exist
    book = xlwt.Workbook(encoding="utf-8")
    book.add_sheet("fights overview")
    book.save(args.xls_output_filename)

    
    #Create Tid file header to support drag and drop onto html page
    myDate = datetime.datetime.now()

    myprint(output, 'created: '+myDate.strftime("%Y%m%d%H%M%S"))
    myprint(output, 'creator: Drevarr ')
    myprint(output, 'curTab: Overview')
    myprint(output, 'tags: Logs [['+myDate.strftime("%Y")+'-'+myDate.strftime("%m")+' Log Reviews]]')
    myprint(output, 'title: '+myDate.strftime("%Y%m%d")+' WvW Log Review\n')
    #End Tid file header

    #JEL-Tweaked to output TW5 formatting (https://drevarr.github.io/FluxCapacity.html)
    print_string = "__''Flux Capacity Node Farmers - WVW Log Review''__\n"
    myprint(output, print_string)

    # print overall stats
    overall_squad_stats = get_overall_squad_stats(fights, config)
    overall_raid_stats = get_overall_raid_stats(fights)
    total_fight_duration = print_total_squad_stats(fights, overall_squad_stats, overall_raid_stats, found_healing, found_barrier, config, output)

    #Start nav_bar_menu for TW5
    Nav_Bar_Items= ('<$button set="!!curTab" setTo="Overview" selectedClass="" class="btn btn-sm btn-dark" style=""> Session Overview </$button>',
                    '<$button set="!!curTab" setTo="Squad Composition" selectedClass="" class="btn btn-sm btn-dark" style=""> Squad Composition </$button>',
                    '<$button set="!!curTab" setTo="Deaths" selectedClass="" class="btn btn-sm btn-dark" style=""> Deaths </$button>',
                    '<$button set="!!curTab" setTo="Damage" selectedClass="" class="btn btn-sm btn-dark" style=""> Damage </$button>',
                    '<$button set="!!curTab" setTo="Damage Taken" selectedClass="" class="btn btn-sm btn-dark" style=""> Damage Taken</$button>',
                    '<$button set="!!curTab" setTo="Boon Strips" selectedClass="" class="btn btn-sm btn-dark" style=""> Boon Strips </$button>',
                    '<$button set="!!curTab" setTo="Condition Cleanses" selectedClass="" class="btn btn-sm btn-dark" style=""> Condition Cleanses</$button>',
                    '<$button set="!!curTab" setTo="Superspeed" selectedClass="" class="btn btn-sm btn-dark" style=""> Superspeed </$button>',
                    '<$button set="!!curTab" setTo="Distance to Tag" selectedClass="" class="btn btn-sm btn-dark" style=""> Distance to Tag </$button>',
                    '<$button set="!!curTab" setTo="Stability" selectedClass="" class="btn btn-sm btn-dark" style=""> Stability </$button>',
                    '<$button set="!!curTab" setTo="Protection" selectedClass="" class="btn btn-sm btn-dark" style=""> Protection </$button>',
                    '<$button set="!!curTab" setTo="Aegis" selectedClass="" class="btn btn-sm btn-dark" style=""> Aegis </$button>',
                    '<$button set="!!curTab" setTo="Might" selectedClass="" class="btn btn-sm btn-dark" style=""> Might </$button>',
                    '<$button set="!!curTab" setTo="Fury" selectedClass="" class="btn btn-sm btn-dark" style=""> Fury </$button>',
                    '<$button set="!!curTab" setTo="Resistance" selectedClass="" class="btn btn-sm btn-dark" style=""> Resistance </$button>',
                    '<$button set="!!curTab" setTo="Resolution" selectedClass="" class="btn btn-sm btn-dark" style=""> Resolution </$button>',
                    '<$button set="!!curTab" setTo="Quickness" selectedClass="" class="btn btn-sm btn-dark" style=""> Quickness </$button>',
                    '<$button set="!!curTab" setTo="Swiftness" selectedClass="" class="btn btn-sm btn-dark" style=""> Swiftness </$button>',
                    '<$button set="!!curTab" setTo="Alacrity" selectedClass="" class="btn btn-sm btn-dark" style=""> Alacrity </$button>',
                    '<$button set="!!curTab" setTo="Vigor" selectedClass="" class="btn btn-sm btn-dark" style=""> Vigor </$button>',
                    '<$button set="!!curTab" setTo="Regeneration" selectedClass="" class="btn btn-sm btn-dark" style=""> Regeneration </$button>',
                    '<$button set="!!curTab" setTo="Support" selectedClass="" class="btn btn-sm btn-dark" style=""> Support Players </$button>'
    )
    for item in Nav_Bar_Items:
        myprint(output, item)
    
    myprint(output, '\n---\n')

    #End nav_bar_menu for TW5

    #Overview reveal
    myprint(output, '<$reveal type="match" state="!!curTab" text="Overview">')
    myprint(output, '\n!!!OVERVIEW\n')

    print_fights_overview(fights, overall_squad_stats, overall_raid_stats, config, output)

    #End reveal
    myprint(output, '</$reveal>')

    write_fights_overview_xls(fights, overall_squad_stats, overall_raid_stats, config, args.xls_output_filename)
    
    #Move Squad Composition here so it is first under the fight summaries
    #Squad Composition Testing
    myprint(output, '<$reveal type="match" state="!!curTab" text="Squad Composition">')    
    myprint(output, '\n!!!SQUAD COMPOSITION\n')    
    sort_order = ['Firebrand', 'Scrapper', 'Spellbreaker', "Herald", "Chronomancer", "Reaper", "Scourge", "Dragonhunter", "Guardian", "Elementalist", "Tempest", "Revenant", "Weaver", "Willbender", "Renegade", "Vindicator", "Warrior", "Berserker", "Bladesworn", "Engineer", "Holosmith", "Mechanist", "Ranger", "Druid", "Soulbeast", "Untamed", "Thief", "Daredevil", "Deadeye", "Specter", "Catalyst", "Mesmer", "Mirage", "Virtuoso", "Necromancer", "Harbinger"]

    output_string = ""

    for fight in squad_comp:
        output_string1 = "\n|thead-dark|k\n"
        output_string2 = ""
        output_string1 += "|Fight |"
        output_string2 += "|"+str(fight+1)
        for prof in sort_order:
            if prof in squad_comp[fight]:
                output_string1 += " {{"+str(prof)+"}} |"
                output_string2 += " | "+str(squad_comp[fight][prof])
                
        output_string1 += "h"
        output_string2 += " |\n"
        
        myprint(output, output_string1)
        myprint(output, output_string2)

    #end reveal
    print_string = "</$reveal>\n"
    myprint(output, print_string)     


    # end Squad Composition insert


    # print top x players for all stats. If less then x
    # players, print all. If x-th place doubled, print all with the
    # same amount of top x achieved.
    num_used_fights = overall_raid_stats['num_used_fights']

    top_total_stat_players = {key: list() for key in config.stats_to_compute}
    top_consistent_stat_players = {key: list() for key in config.stats_to_compute}
    top_average_stat_players = {key: list() for key in config.stats_to_compute}
    top_percentage_stat_players = {key: list() for key in config.stats_to_compute}
    top_late_players = {key: list() for key in config.stats_to_compute}
    top_jack_of_all_trades_players = {key: list() for key in config.stats_to_compute}    
    
    #JEL-Tweaked to output TW5 formatting (https://drevarr.github.io/FluxCapacity.html)
    for stat in config.stats_to_compute:
        if (stat == 'heal' and not found_healing) or (stat == 'barrier' and not found_barrier):
            continue

        #JEL-Tweaked to output TW5 output to maintain formatted table and slider (https://drevarr.github.io/FluxCapacity.html)
        myprint(output,'<$reveal type="match" state="!!curTab" text="'+config.stat_names[stat]+'">')
        myprint(output, "\n!!!"+config.stat_names[stat].upper()+"\n")

        if stat == 'dist':
            myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
            top_consistent_stat_players[stat] = get_top_players(players, config, stat, StatType.CONSISTENT)
            top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
            top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)            
            top_percentage_stat_players[stat],comparison_val = get_and_write_sorted_top_percentage(players, config, num_used_fights, stat, output, StatType.PERCENTAGE, top_consistent_stat_players[stat])
            myprint(output, '\n</div>\n</div>\n')
        elif stat == 'dmg_taken':
            myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
            top_consistent_stat_players[stat] = get_top_players(players, config, stat, StatType.CONSISTENT)
            top_total_stat_players[stat] = get_top_players(players, config, stat, StatType.TOTAL)
            top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
            top_average_stat_players[stat] = get_and_write_sorted_average(players, config, num_used_fights, stat, output)
            myprint(output, '\n</div>\n</div>\n')
        else:
            myprint(output, '\n<div class="flex-row">\n    <div class="flex-col border">\n')
            top_consistent_stat_players[stat] = get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output)
            myprint(output, '\n</div>\n    <div class="flex-col border">\n')
            top_total_stat_players[stat] = get_and_write_sorted_total(players, config, total_fight_duration, stat, output)
            myprint(output, '\n</div>\n</div>\n')
            top_average_stat_players[stat] = get_top_players(players, config, stat, StatType.AVERAGE)
            top_percentage_stat_players[stat],comparison_val = get_top_percentage_players(players, config, stat, StatType.PERCENTAGE, num_used_fights, top_consistent_stat_players[stat], top_total_stat_players[stat], list(), list())
        
        #JEL-Tweaked to output TW5 output to maintain formatted table and slider (https://drevarr.github.io/FluxCapacity.html)
        myprint(output, "</$reveal>\n")

        write_to_json(overall_raid_stats, overall_squad_stats, fights, players, top_total_stat_players, top_average_stat_players, top_consistent_stat_players, top_percentage_stat_players, top_late_players, top_jack_of_all_trades_players, args.json_output_filename)

    #print table of accounts that fielded support characters
    myprint(output,'<$reveal type="match" state="!!curTab" text="Support">')
    myprint(output, "\n")
    # print table header
    print_string = "|thead-dark table-hover sortable|k"    
    myprint(output, print_string)
    print_string = "|!Account |!Name |!Profession | !Fights| !Duration|!Support |!Guild Status |h"
    myprint(output, print_string)    

    for stat in config.stats_to_compute:
        if (stat == 'rips' or stat == 'cleanses' or stat == 'stability'):
            write_support_players(players, top_total_stat_players[stat], stat, output)

    myprint(output, "</$reveal>\n")

    supportCount=0
    for stat in config.stats_to_compute:
        if stat == 'dist':
            write_stats_xls(players, top_percentage_stat_players[stat], stat, args.xls_output_filename)
        elif stat == 'dmg_taken':
            write_stats_xls(players, top_average_stat_players[stat], stat, args.xls_output_filename)
        elif stat == 'heal' and found_healing:
            write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)            
        elif stat == 'barrier' and found_barrier:
            write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)
        elif stat == 'deaths':
            write_stats_xls(players, top_consistent_stat_players[stat], stat, args.xls_output_filename)
        else:
            write_stats_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename)
            if stat == 'rips' or stat == 'cleanses' or stat == 'stability':
                supportCount = write_support_xls(players, top_total_stat_players[stat], stat, args.xls_output_filename, supportCount)