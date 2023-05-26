#!/usr/bin/env python3

from os import listdir
import argparse
import os.path
import sys
import importlib
import json
import xlwt
import datetime
import gzip

from TW5_parse_top_stats_tools import fill_config, reset_globals, get_stats_from_fight_json, get_stat_from_player_json, get_buff_ids_from_json, get_combat_time_breakpoints, sum_breakpoints, BuffGenerationType

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='This reads a set of arcdps reports in xml format and generates top stats.')
	parser.add_argument('input_directory', help='Directory containing .xml or .json files from arcdps reports')
	parser.add_argument('-c', '--config_file', dest="config_file", help="Config file with all the settings", default="TW5_parser_config_detailed")
	parser.add_argument('-a', '--anonymized', dest="anonymize", help="Create an anonymized version of the top stats. All account and character names will be replaced.", default=False, action='store_true')
	args = parser.parse_args()

	if not os.path.isdir(args.input_directory):
		print("Directory ",args.input_directory," is not a directory or does not exist!")
		sys.exit()

	log = open(args.input_directory+"/log_detailed.txt", "w")

	parser_config = importlib.import_module("parser_configs."+args.config_file , package=None) 
	config = fill_config(parser_config)
	
	# create xls file if it doesn't exist
	book = xlwt.Workbook(encoding="utf-8")
	sheet1 = book.add_sheet("Player Stats")

	sheet1.write(0, 0, "Account")
	sheet1.write(0, 1, "Name")
	sheet1.write(0, 2, "Profession")
	sheet1.write(0, 3, "Role")
	sheet1.write(0, 4, "Rally Num")
	sheet1.write(0, 5, "Fight Num")
	sheet1.write(0, 6, "Date")
	sheet1.write(0, 7, "Start Time")
	sheet1.write(0, 8, "End Time")
	sheet1.write(0, 9, "Num Allies")
	sheet1.write(0, 10, "Num Party Members")
	sheet1.write(0, 11, "Num Enemies")
	sheet1.write(0, 12, "Group")
	sheet1.write(0, 13, "Duration")
	sheet1.write(0, 14, "Combat time")
	sheet1.write(0, 15, "Damage")
	sheet1.write(0, 16, "Power Damage")
	sheet1.write(0, 17, "Condi Damage")
	sheet1.write(0, 18, "Crit Perc")
	sheet1.write(0, 19, "Flanking Perc")
	sheet1.write(0, 20, "Glancing Perc")
	sheet1.write(0, 21, "Blind Num")
	sheet1.write(0, 22, "Interrupt Num")
	sheet1.write(0, 23, "Invulnerable Num")
	sheet1.write(0, 24, "Evaded Num")
	sheet1.write(0, 25, "Blocked Num")

	stats_to_compute = ['downs', 'kills', 'res', 'deaths', 'dmg_taken', 'barrierDamage', 'dist',  'swaps', 'rips', 'cleanses', 'barrier']
	for i,stat in enumerate(stats_to_compute):
		sheet1.write(0, 26+i, config.stat_names[stat])
	
	sheet1.write(0, 37, 'Total Healing')
	sheet1.write(0, 38, 'Power Healing')
	sheet1.write(0, 39, 'Conversion Healing')
	sheet1.write(0, 40, 'Hybrid Healing')
	sheet1.write(0, 41, 'Group Total Healing')
	sheet1.write(0, 42, 'Group Power Healing')
	sheet1.write(0, 43, 'Group Conversion Healing')
	sheet1.write(0, 44, 'Group Hybrid Healing')

	uptime_Order = ['protection',  'aegis',  'fury',  'resistance',  'resolution',  'quickness',  'swiftness',  'alacrity',  'vigor',  'regeneration']
	stacking_uptime_Order =  ['stability',  'might']

	for i,stat in enumerate(uptime_Order + stacking_uptime_Order):
		sheet1.write(0, 45+i, stat.capitalize()+" Squad Gen")

	for i,stat in enumerate(uptime_Order + stacking_uptime_Order):
		sheet1.write(0, 57+i, stat.capitalize()+" Group Gen")

	for i,stat in enumerate(uptime_Order + stacking_uptime_Order):
		sheet1.write(0, 69+i, stat.capitalize()+" Self Gen")

	for i,stat in enumerate(uptime_Order):
		sheet1.write(0, 81+i, stat.capitalize()+" Uptime")

	sheet1.write(0, 91, "Stab Avg Stacks")
	sheet1.write(0, 92, "Stab 1+ Stacks")
	sheet1.write(0, 93, "Stab 2+ Stacks")
	sheet1.write(0, 94, "Stab 5+ Stacks")

	sheet1.write(0, 95, "Might Avg Stacks")
	sheet1.write(0, 96, "Might 10+ Stacks")
	sheet1.write(0, 97, "Might 15+ Stacks")
	sheet1.write(0, 98, "Might 20+ Stacks")
	sheet1.write(0, 99, "Might 25 Stacks")

	for i,stat in enumerate(uptime_Order + ['stability']):
		sheet1.write(0, 100+i, stat.capitalize()+" Uptime while damaging")

	sheet1.write(0, 111, "Might Avg Stack while damaging")
	sheet1.write(0, 112, "Might 10+ Stacks while damaging")
	sheet1.write(0, 113, "Might 15+ Stacks while damaging")
	sheet1.write(0, 114, "Might 20+ Stacks while damaging")
	sheet1.write(0, 115, "Might 25 Stacks while damaging")

	sheet1.write(0, 116, "Carrion Damage")
	for i in range(1, 21):
		sheet1.write(0, 116 + i, 'Chunk Damage (' + str(i) + ')')
	for i in range(1, 21):
		sheet1.write(0, 136 + i, 'Burst Damage (' + str(i) + ')')
	for i in range(1, 21):
		sheet1.write(0, 156 + i, 'Ch5Ca Burst Damage (' + str(i) + ')')


	# iterating over all fights in directory
	files = listdir(args.input_directory)
	sorted_files = sorted(files)
	rally_num = 1
	fight_num = 1
	last_fight_end_time = None
	row = 1
	for filename in sorted_files:
		file_start, file_extension = os.path.splitext(filename)
		if file_extension not in ['.json', '.gz'] or "top_stats" in file_start:
			continue

		print("parsing "+filename)
		file_path = "".join((args.input_directory,"/",filename))

		if file_extension == '.gz':
			with gzip.open(file_path, mode="r") as f:
				json_data = json.loads(f.read().decode('utf-8'))
		else:
			json_datafile = open(file_path, encoding='utf-8')
			json_data = json.load(json_datafile)

		reset_globals()
		config = fill_config(parser_config)
		get_buff_ids_from_json(json_data, config)
		fight, players_running_healing_addon, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableIn, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats = get_stats_from_fight_json(json_data, config, log)
		
		if fight.skipped:
			continue


		if last_fight_end_time:
			after_last_fight = datetime.datetime.fromisoformat(last_fight_end_time) + datetime.timedelta(hours=2)
			if after_last_fight <  datetime.datetime.fromisoformat(fight.start_time):
				print("Start of a new rally at ", fight.start_time)
				rally_num += 1
				fight_num = 1

		last_fight_end_time = fight.end_time

		party_member_counts = {}
		for player_data in json_data['players']:
			group = player_data['group']
			if group in party_member_counts:
				party_member_counts[group] += 1
			else:
				party_member_counts[group] = 1

		for squadDps_prof_name in DPSStats:
			player = [p for p in json_data["players"] if p['account'] == DPSStats[squadDps_prof_name]['account']][0]
			player_prof_name = "{{"+player['profession']+"}} "+player['name']

			fight_duration = json_data["durationMS"] / 1000
			combat_time = sum_breakpoints(get_combat_time_breakpoints(player)) / 1000
			num_party_members = party_member_counts[player['group']]

			# Calculate healing values
			if player['name'] in players_running_healing_addon and 'extHealingStats' in player:
				total_healing = 0
				total_healing_group = 0
				power_healing = 0
				power_healing_group = 0
				conversion_healing = 0
				conversion_healing_group = 0
				hybrid_healing = 0
				hybrid_healing_group = 0

				allied_healing_1s = player['extHealingStats']['alliedHealing1S']
				allied_power_healing_1s = player['extHealingStats']['alliedHealingPowerHealing1S']
				allied_conversion_healing_1s = player['extHealingStats']['alliedConversionHealingHealing1S']
				allied_hybrid_healing_1s = player['extHealingStats']['alliedHybridHealing1S']
				for index in range(len(json_data['players'])):
					is_same_group = player['group'] == json_data['players'][index]['group']

					player_healing = allied_healing_1s[index][0][-1]
					player_power_healing = allied_power_healing_1s[index][0][-1]
					player_conversion_healing = allied_conversion_healing_1s[index][0][-1]
					player_hybrid_healing = allied_hybrid_healing_1s[index][0][-1]


					total_healing += player_healing
					power_healing += player_power_healing
					conversion_healing += player_conversion_healing
					hybrid_healing += player_hybrid_healing
					if is_same_group:
						total_healing_group += player_healing
						power_healing_group += player_power_healing
						conversion_healing_group += player_conversion_healing
						hybrid_healing_group += player_hybrid_healing
			else:
				# When no healing stats data, set to -1
				total_healing = -1
				total_healing_group = -1
				power_healing = -1
				power_healing_group = -1
				conversion_healing = -1
				conversion_healing_group = -1
				hybrid_healing = -1
				hybrid_healing_group = -1

			sheet1.write(row, 0, DPSStats[squadDps_prof_name]['account'])
			sheet1.write(row, 1, DPSStats[squadDps_prof_name]['name'])
			sheet1.write(row, 2, DPSStats[squadDps_prof_name]['profession'])
			sheet1.write(row, 3, DPSStats[squadDps_prof_name]['role'])
			sheet1.write(row, 4, rally_num)
			sheet1.write(row, 5, fight_num)
			sheet1.write(row, 6, fight.start_time.split()[0])
			sheet1.write(row, 7, fight.start_time.split()[1])
			sheet1.write(row, 8, fight.end_time.split()[1])
			sheet1.write(row, 9, fight.allies)
			sheet1.write(row, 10, num_party_members)
			sheet1.write(row, 11, fight.enemies)
			sheet1.write(row, 12, int(player['group']))
			sheet1.write(row, 13, fight_duration)
			sheet1.write(row, 14, combat_time)
			sheet1.write(row, 15, get_stat_from_player_json(player, players_running_healing_addon, 'dmg', config))
			sheet1.write(row, 16, get_stat_from_player_json(player, players_running_healing_addon, 'Pdmg', config))
			sheet1.write(row, 17, get_stat_from_player_json(player, players_running_healing_addon, 'Cdmg', config))

			if squad_offensive[player_prof_name]['stats']['critableDirectDamageCount'] > 0:
				sheet1.write(row, 18, squad_offensive[player_prof_name]['stats']['criticalRate'] / squad_offensive[player_prof_name]['stats']['critableDirectDamageCount'])
			else:
				sheet1.write(row, 18, 0)

			if squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'] > 0:
				sheet1.write(row, 19, squad_offensive[player_prof_name]['stats']['flankingRate'] / squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'])
				sheet1.write(row, 20, squad_offensive[player_prof_name]['stats']['glanceRate'] / squad_offensive[player_prof_name]['stats']['connectedDirectDamageCount'])
			else:
				sheet1.write(row, 19, 0)
				sheet1.write(row, 20, 0)

			sheet1.write(row, 21, squad_offensive[player_prof_name]['stats']['missed'])
			sheet1.write(row, 22, squad_offensive[player_prof_name]['stats']['interrupts'])
			sheet1.write(row, 23, squad_offensive[player_prof_name]['stats']['invulned'])
			sheet1.write(row, 24, squad_offensive[player_prof_name]['stats']['evaded'])
			sheet1.write(row, 25, squad_offensive[player_prof_name]['stats']['blocked'])

			for i,stat in enumerate(stats_to_compute):
				sheet1.write(row, 26+i, get_stat_from_player_json(player, players_running_healing_addon, stat, config))
			
			sheet1.write(row, 37, total_healing)
			sheet1.write(row, 38, power_healing)
			sheet1.write(row, 39, conversion_healing)
			sheet1.write(row, 40, hybrid_healing)
			sheet1.write(row, 41, total_healing_group)
			sheet1.write(row, 42, power_healing_group)
			sheet1.write(row, 43, conversion_healing_group)
			sheet1.write(row, 44, hybrid_healing_group)

			for i,stat in enumerate(uptime_Order):
				sheet1.write(row, 45+i, (get_stat_from_player_json(player, players_running_healing_addon, stat, config)/100)*fight_duration*(fight.allies - 1))
			for i,stat in enumerate(stacking_uptime_Order):
				sheet1.write(row, 55+i, get_stat_from_player_json(player, players_running_healing_addon, stat, config)*fight_duration*(fight.allies - 1))

			for i,stat in enumerate(uptime_Order):
				sheet1.write(row, 57+i, (get_stat_from_player_json(player, players_running_healing_addon, stat, config, False, BuffGenerationType.GROUP)/100)*fight_duration*(num_party_members - 1))
			for i,stat in enumerate(stacking_uptime_Order):
				sheet1.write(row, 67+i, get_stat_from_player_json(player, players_running_healing_addon, stat, config, False, BuffGenerationType.GROUP)*fight_duration*(num_party_members - 1))

			for i,stat in enumerate(uptime_Order):
				sheet1.write(row, 69+i, (get_stat_from_player_json(player, players_running_healing_addon, stat, config, False, BuffGenerationType.SELF)/100)*fight_duration)
			for i,stat in enumerate(stacking_uptime_Order):
				sheet1.write(row, 79+i, get_stat_from_player_json(player, players_running_healing_addon, stat, config, False, BuffGenerationType.SELF)*fight_duration)
		
			uptime_duration = uptime_Table[player_prof_name]['duration']
			for i,stat in enumerate(uptime_Order):
				if stat in uptime_Table[player_prof_name]:
					buff_Time = uptime_Table[player_prof_name][stat]
					sheet1.write(row, 81+i, buff_Time / uptime_duration)
				else:
					sheet1.write(row, 81+i, 0.00)

			stability_stacks = stacking_uptime_Table[squadDps_prof_name]['stability']
			stability_stacks_fight_time = (stacking_uptime_Table[squadDps_prof_name]['duration_stability'] / 1000) or 1

			might_stacks = stacking_uptime_Table[squadDps_prof_name]['might']
			might_stacks_fight_time = (stacking_uptime_Table[squadDps_prof_name]['duration_might'] / 1000) or 1
			
			sheet1.write(row, 91, sum(stack_num * stability_stacks[stack_num] for stack_num in range(1, 26)) / (stability_stacks_fight_time * 1000))
			sheet1.write(row, 92, (1.0 - (stability_stacks[0] / (stability_stacks_fight_time * 1000))))
			sheet1.write(row, 93, sum(stability_stacks[i] for i in range(2,26)) / (stability_stacks_fight_time * 1000))
			sheet1.write(row, 94, sum(stability_stacks[i] for i in range(5,26)) / (stability_stacks_fight_time * 1000))

			sheet1.write(row, 95, sum(stack_num * might_stacks[stack_num] for stack_num in range(1, 26)) / (might_stacks_fight_time * 1000))
			sheet1.write(row, 96, sum(might_stacks[i] for i in range(10,26)) / (might_stacks_fight_time * 1000))
			sheet1.write(row, 97, sum(might_stacks[i] for i in range(15,26)) / (might_stacks_fight_time * 1000))
			sheet1.write(row, 98, sum(might_stacks[i] for i in range(20,26)) / (might_stacks_fight_time * 1000))
			sheet1.write(row, 99, might_stacks[25] / (might_stacks_fight_time * 1000))

			total_damage = DPSStats[squadDps_prof_name]["Damage_Total"] or 1

			for i,stat in enumerate(uptime_Order + ['stability']):
				sheet1.write(row, 100+i, stacking_uptime_Table[squadDps_prof_name]['damage_with_'+stat][1] / total_damage)
			
			damage_with_might = stacking_uptime_Table[squadDps_prof_name]['damage_with_might']

			sheet1.write(row, 111, sum(stack_num * damage_with_might[stack_num] for stack_num in range(1, 26)) / total_damage)
			sheet1.write(row, 112, sum(damage_with_might[i] for i in range(10,26)) / total_damage)
			sheet1.write(row, 113, sum(damage_with_might[i] for i in range(15,26)) / total_damage)
			sheet1.write(row, 114, sum(damage_with_might[i] for i in range(20,26)) / total_damage)
			sheet1.write(row, 115, damage_with_might[25] / total_damage)			

			sheet1.write(row, 116, DPSStats[squadDps_prof_name]['Carrion_Damage'])

			for i in range(1, 21):
				sheet1.write(row, 116 + i, DPSStats[squadDps_prof_name]['Chunk_Damage'][i])

			for i in range(1, 21):
				sheet1.write(row, 136 + i, DPSStats[squadDps_prof_name]['Burst_Damage'][i])

			for i in range(1, 21):
				sheet1.write(row, 156 + i, DPSStats[squadDps_prof_name]['Ch5Ca_Burst_Damage'][i])

			row += 1
		
		fight_num += 1

	book.save(args.input_directory+"/TW5_top_stats_per_fight.xls")