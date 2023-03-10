#!/usr/bin/env python3

#    parse_top_stats_tools.py contains tools for computing top stats in arcdps logs as parsed by Elite Insights.
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


from cgi import test
from dataclasses import dataclass,field
import os.path
from os import listdir
import sys
#import xml.etree.ElementTree as ET
from enum import Enum
import importlib
import xlrd
from xlutils.copy import copy
import json
import jsons
import math
import requests
import datetime
import gzip
from collections import OrderedDict

from GW2_Color_Scheme import ProfessionColor

try:
    import Guild_Data
except ImportError:
    Guild_Data = None


debug = False # enable / disable debug output

class StatType(Enum):
	TOTAL = 1
	CONSISTENT = 2
	AVERAGE = 3
	LATE_PERCENTAGE = 4
	SWAPPED_PERCENTAGE = 5
	PERCENTAGE = 6

class BuffGenerationType(Enum):
	SQUAD = 1
	GROUP = 2
	OFFGROUP = 3
	SELF = 4
	

# This class stores information about a player. Note that a different profession will be treated as a new player / character.
@dataclass
class Player:
	account: str                        # account name
	name: str                           # character name
	profession: str                     # profession name
	num_fights_present: int = 0         # the number of fight the player was involved in 
	num_enemies_present: int = 0        # the number of fight the player was involved in
	num_allies_supported: int = 0       # the number of fight the player was involved in
	num_allies_group_supported: int = 0  # the number of fight the player was involved in
	attendance_percentage: float = 0.   # the percentage of fights the player was involved in out of all fights
	duration_fights_present: int = 0    # the total duration of all fights the player was involved in, in s
	duration_active: int = 0            # the total duration a player was active (alive or down)
	duration_in_combat: int = 0         # the total duration a player was in combat (taking/dealing dmg)    
	swapped_build: bool = False         # a different player character or specialization with this account name was in some of the fights 

	# fields for all stats defined in config
	consistency_stats: dict = field(default_factory=dict)     # how many times did this player get into top for each stat?
	total_stats: dict = field(default_factory=dict)           # what's the total value for this player for each stat?
	total_stats_group: dict = field(default_factory=dict)     # what's the total value for this player for each stat for their group?
	total_stats_self: dict = field(default_factory=dict)      # what's the total value for this player for each stat for self?
	average_stats: dict = field(default_factory=dict)         # what's the average stat per second for this player? (exception: deaths are per minute)
	portion_top_stats: dict = field(default_factory=dict)     # what percentage of fights did this player get into top for each stat, in relation to the number of fights they were involved in?
	stats_per_fight: list = field(default_factory=list)       # what's the value of each stat for this player in each fight?

	#fields for wt dps per enemy
	wt_dps_enemies: list = field(default_factory=list)      # list of enemies present by fight
	wt_dps_duration: list = field(default_factory=list)     # list of enemies present by fight
	wt_dps_damage: list = field(default_factory=list)       # list of enemies present by fight

	def initialize(self, config):
		self.total_stats = {key: 0 for key in config.stats_to_compute}
		self.total_stats_group = {key: 0 for key in config.stats_to_compute}
		self.total_stats_self = {key: 0 for key in config.stats_to_compute}
		self.average_stats = {key: 0 for key in config.stats_to_compute}        
		self.consistency_stats = {key: 0 for key in config.stats_to_compute}
		self.portion_top_stats = {key: 0 for key in config.stats_to_compute}


# This class stores information about a fight
@dataclass
class Fight:
	skipped: bool = False
	duration: int = 0
	total_stats: dict = field(default_factory=dict) # what's the over total value for the whole squad for each stat in this fight?
	enemies: int = 0
	allies: int = 0
	kills: int = 0
	start_time: str = ""
	enemy_squad: dict = field(default_factory=dict) #profession and count of enemies
	enemy_Dps: dict = field(default_factory=dict) #enemy name and amount of damage output
	squad_Dps: dict = field(default_factory=dict) #squad player name and amount of damage output
	enemy_skill_dmg: dict = field(default_factory=dict) #enemy skill_name and amount of damage output
	squad_skill_dmg: dict = field(default_factory=dict) #squad skill_name and amount of damage output
	squad_spike_dmg: dict = field(default_factory=dict) #squad skill_name and amount of damage output

	
	
# This class stores the configuration for running the top stats.
@dataclass
class Config:
	num_players_listed: dict = field(default_factory=dict)          # How many players will be listed who achieved top stats most often for each stat?
	num_players_considered_top_percentage: float = 0.  # % of players considered to be "top" in each fight for each stat?

	min_attendance_portion_for_percentage: float = 0.  # For what portion of all fights does a player need to be there to be considered for "percentage" awards?
	min_attendance_portion_for_late: float = 0.        # For what portion of all fights does a player need to be there to be considered for "late but great" awards?     
	min_attendance_portion_for_buildswap: float = 0.   # For what portion of all fights does a player need to be there to be considered for "jack of all trades" awards?
	min_attendance_percentage_for_average: float = 0.  # For what percentage of all fights does a player need to be there to be considered for "jack of all trades" awards?     

	portion_of_top_for_total: float = 0.         # What portion of the top total player stat does someone need to reach to be considered for total awards?
	portion_of_topDamage_for_total: float = 0.         # What portion of the top total player stat does someone need to reach to be considered for total awards?	
	portion_of_top_for_consistent: float = 0.    # What portion of the total stat of the top consistent player does someone need to reach to be considered for consistency awards?
	portion_of_top_for_percentage: float = 0.    # What portion of the consistency stat of the top consistent player does someone need to reach to be considered for percentage awards?    
	portion_of_top_for_late: float = 0.          # What portion of the percentage the top consistent player reached top does someone need to reach to be considered for late but great awards?
	portion_of_top_for_buildswap: float = 0.     # What portion of the percentage the top consistent player reached top does someone need to reach to be considered for jack of all trades awards?

	min_allied_players: int = 0   # minimum number of allied players to consider a fight in the stats
	min_fight_duration: int = 0   # minimum duration of a fight to be considered in the stats
	min_enemy_players: int = 0    # minimum number of enemies to consider a fight in the stats

	summary_title: str = ""
	summary_creator: str = ""

	charts: bool = False	# produce charts for stats_to_compute

	stat_names: dict = field(default_factory=dict)
	profession_abbreviations: dict = field(default_factory=dict)

	empty_stats: dict = field(default_factory=dict)
	stats_to_compute: list = field(default_factory=list)
	aurasIn_to_compute: list = field(default_factory=list)
	aurasOut_to_compute: list = field(default_factory=list)
	defenses_to_compute: list = field(default_factory=list)

	buff_ids: dict = field(default_factory=dict)
	buffs_stacking_duration: list = field(default_factory=list)
	buffs_stacking_intensity: list = field(default_factory=list)
	buff_abbrev: dict = field(default_factory=dict)
	condition_ids: dict = field(default_factory=dict)
	auras_ids: dict = field(default_factory=dict)


#Stats to exlucde from overview summary
exclude_Stat = ["iol", "dist", "res", "Cdmg", "Pdmg",  "kills", "downs", 'downed', "HiS", "stealth", "superspeed", "swaps", "barrierDamage", "dodges", "evades", "blocks", "invulns", 'hitsMissed', 'interupted', 'fireOut', 'shockingOut', 'frostOut', 'magneticOut', 'lightOut', 'darkOut', 'chaosOut', 'ripsIn', 'ripsTime', 'cleansesIn', 'cleansesTime', 'downContrib']

#Control Effects Tracking
squad_offensive = {}
squad_Control = {} 
enemy_Control = {} 
enemy_Control_Player = {} 

#Spike Damage Tracking
squad_damage_output = {}

#Downed Healing from Instant Revive skills
downed_Healing = {}

#Aura Tracking 
auras_TableOut = {}

#Uptime Tracking
uptime_Table = {}
uptime_Buff_Ids = {1122: 'stability', 717: 'protection', 743: 'aegis', 740: 'might', 725: 'fury', 26980: 'resistance', 873: 'resolution', 1187: 'quickness', 719: 'swiftness', 30328: 'alacrity', 726: 'vigor', 718: 'regeneration'}
#uptime_Buff_Names = { 'stability': 1122,  'protection': 717,  'aegis': 743,  'might': 740,  'fury': 725,  'resistance': 26980,  'resolution': 873,  'quickness': 1187,  'swiftness': 719,  'alacrity': 30328,  'vigor': 726,  'regeneration': 718}

#Stacking Buffs Tracking
stacking_uptime_Table = {}

#Personal Buff Tracking
buffs_personal = {}

#Skill Dictionary from all Fights
skill_Dict = {}

#Calculate On Tag Death Variables
On_Tag = 600
Run_Back = 5000
Death_OnTag = {}

#Collect Account Attendance Data
Attendance = {}

#Collect DPS Box Plot Data
DPS_List = {}
DPS_List['acct'] = {}
DPS_List['name'] = {}
DPS_List['prof_name'] = {}
DPS_List['prof'] = {}

#Collect CPS Box Plot Data
CPS_List = {}
CPS_List['acct'] = {}
CPS_List['name'] = {}
CPS_List['prof_name'] = {}
CPS_List['prof'] = {}

#Collect CPS Box Plot Data
SPS_List = {}
SPS_List['acct'] = {}
SPS_List['name'] = {}
SPS_List['prof_name'] = {}
SPS_List['prof'] = {}

#Collect CPS Box Plot Data
HPS_List = {}
HPS_List['acct'] = {}
HPS_List['name'] = {}
HPS_List['prof_name'] = {}
HPS_List['prof'] = {}

#Calculate DPSStats Variables
DPSStats = {}

#Collect MOA Info
MOA_Targets = {}
MOA_Casters = {}

#fetch Guild Data and Check Guild Status function
#members: Dict[str, Any] = {}
members: dict = field(default_factory=dict) 
API_response = ""

if Guild_Data:
	if type(Guild_Data.Guild_ID) == dict:
		print("Guild Keys Available: ", Guild_Data.Guild_ID.keys())
		inputGuild = input('What Guild key for this session?\n')
		Guild_ID = Guild_Data.Guild_ID[inputGuild]
		API_Key = Guild_Data.API_Key[inputGuild]
		print("-----=====:Guild ID:", Guild_ID)
		print("-----=====:Guild API:", API_Key)
		api_url = "https://api.guildwars2.com/v2/guild/"+Guild_ID+"/members?access_token="+API_Key
	else:
		Guild_ID = Guild_Data.Guild_ID
		API_Key = Guild_Data.API_Key
		api_url = "https://api.guildwars2.com/v2/guild/"+Guild_ID+"/members?access_token="+API_Key
	response = requests.get(api_url)
	members = json.loads(response.text)
	print("response code: "+str(response.status_code))
	API_response = response.status_code
else:
	members = {}
	API_response = " "


def findMember(json_object, name):
	if API_response == requests.codes.ok:
		guildStatus = "--==Non Member==--"
		for dict in json_object:
			if dict['name'] == name:
				guildStatus = dict['rank']
		return guildStatus
	else:
		guildStatus = ""
		return guildStatus
# End fetch Guild Data and Check Guild Status

#define subtype based on consumables

#consumable dictionaries
Heal_Utility = {
    53374: "Potent Lucent Oil", 
    53304: "Enhanced Lucent Oil", 
    21827: "Toxic Maintenance Oil", 
    34187: "Peppermint Oil", 
    38605: "Magnanimous Maintenance Oil",
    25879: "Bountiful Maintenance Oil",
    9968: "Master Maintenance Oil"
    }
Heal_Food = {
    57276: "Bowl of Spiced Fruit Salad",
    57100: "Bowl of Fruit Salad with Mint Garnish",
    26529: "Delicious Rice Ball"
    }
Cele_Food = {
    57165: "Spherified Peppercorn-Spiced Oyster Soup",
    57374: "Spherified Clove-Spiced Oyster Soup",
    57037: "Spherified Sesame Oyster Soup",
    57201: "Spherified Oyster Soup with Mint Garnish",
    19451: "Dragon's Revelry Starcake"
    }
DPS_Food= {
	57051: "Peppercorn-Crusted Sous-Vide Steak",
	57244: "Cilantro Lime Sous-Vide Steak",
	57260: "Plate of Peppercorn-Spiced Coq Au Vin"
	}
DPS_Utility = {
	9963: "Superior Sharpening Stone",
	34657: "Compact Hardened Sharpening Stone",
	25882: "Furious Sharpening Stone",
	33297: "Writ of Masterful Strength"
	}

def find_sub_type(player, fightTime):
	supportProf = ["Tempest", "Scrapper", "Druid", "Chronomancer", "Vindicator", "Firebrand", "Spectre", "Spellbreaker", "Willbender", "Guardian"]
	if player['profession'] not in supportProf:

		playerDamage = 0
		playerPowerDamage = 0
		playerCondiDamage = 0
		for target in player['dpsTargets']:
			playerDamage += target[0]['damage']
			playerPowerDamage += target[0]['powerDamage']
			playerCondiDamage += target[0]['condiDamage']

		if 'consumables' in player:
			for item in player['consumables']:
				if item['id'] in Cele_Food:
					return "Cele"
				
		# If a player is predominantly condi damage
		if playerCondiDamage > playerPowerDamage:
			return "Condi"
		
		# assume DPS on a nonSupport profession
		else:
			return "Dps"
		
	criticalCount = sum([stats[0]['criticalRate'] for stats in player['statsTargets']])
	critableCount = sum([stats[0]['critableDirectDamageCount'] for stats in player['statsTargets']])
	critPercent = criticalCount / critableCount if critableCount else 0

	# Only healers should have a crit % lower than 40%
	if critPercent <= 0.4:
		return "Support"

	#adjusted consumable search since food and utility can reside in any consumable slot
	if 'consumables' in player:
		for item in player['consumables']:
			if item['id'] in Cele_Food:
				return "Cele"
		for item in player['consumables']:			
			if item['id'] in Heal_Food or item['id'] in Heal_Utility:
				return "Support" 
		for item in player['consumables']:			
			if item['id'] in DPS_Food or item['id'] in DPS_Utility:
				return "Dps"

	# If all other detection fails, fallback to assuming DPS like before
	return "Dps"
#end define subtype based on consumables


# prints output_string to the console and the output_file, with a linebreak at the end
def myprint(output_file, output_string):
	print(output_string)
	output_file.write(output_string+"\n")


# JEL - format a number with commas every thousand
def my_value(number):
	return ("{:,}".format(number))


# fills a Config with the given input    
def fill_config(config_input):
	config = Config()
	config.num_players_listed = config_input.num_players_listed
	config.num_players_considered_top = config_input.num_players_considered_top_percentage/100
	
	config.player_sorting_stat_type = config_input.player_sorting_stat_type or 'total'

	config.min_attendance_portion_for_percentage = config_input.attendance_percentage_for_percentage/100.
	config.min_attendance_portion_for_late = config_input.attendance_percentage_for_late/100.    
	config.min_attendance_portion_for_buildswap = config_input.attendance_percentage_for_buildswap/100.
	config.min_attendance_percentage_for_average = config_input.attendance_percentage_for_average
	config.min_attendance_percentage_for_top = config_input.attendance_percentage_for_top

	config.portion_of_top_for_consistent = config_input.percentage_of_top_for_consistent/100.
	config.portion_of_top_for_total = config_input.percentage_of_top_for_total/100.
	config.portion_of_topDamage_for_total = config_input.percentage_of_topDamage_for_total/100.	
	config.portion_of_top_for_average = config_input.percentage_of_top_for_average/100.	
	config.portion_of_top_for_percentage = config_input.percentage_of_top_for_percentage/100.
	config.portion_of_top_for_late = config_input.percentage_of_top_for_late/100.    
	config.portion_of_top_for_buildswap = config_input.percentage_of_top_for_buildswap/100.

	config.min_allied_players = config_input.min_allied_players
	config.min_fight_duration = config_input.min_fight_duration
	config.min_enemy_players = config_input.min_enemy_players

	config.summary_title = config_input.summary_title
	config.summary_creator = config_input.summary_creator

	config.stat_names = config_input.stat_names
	config.profession_abbreviations = config_input.profession_abbreviations

	config.stats_to_compute = config_input.stats_to_compute
	config.aurasIn_to_compute = config_input.aurasIn_to_compute
	config.aurasOut_to_compute = config_input.aurasOut_to_compute
	config.defenses_to_compute = config_input.defenses_to_compute
	config.empty_stats = {stat: -1 for stat in config.stats_to_compute}
	config.empty_stats['time_active'] = -1
	config.empty_stats['time_in_combat'] = -1

	config.buff_abbrev["Stability"] = 'stability'
	config.buff_abbrev["Protection"] = 'protection'
	config.buff_abbrev["Aegis"] = 'aegis'
	config.buff_abbrev["Might"] = 'might'
	config.buff_abbrev["Fury"] = 'fury'
	config.buff_abbrev["Superspeed"] = 'superspeed'
	config.buff_abbrev["Stealth"] = 'stealth'
	config.buff_abbrev["Hide in Shadows"] = 'HiS'
	config.buff_abbrev["Regeneration"] = 'regeneration'
	config.buff_abbrev["Resistance"] = 'resistance'
	config.buff_abbrev["Resolution"] = 'resolution'
	config.buff_abbrev["Quickness"] = 'quickness'
	config.buff_abbrev["Swiftness"] = 'swiftness'
	config.buff_abbrev["Alacrity"] = 'alacrity'
	config.buff_abbrev["Vigor"] = 'vigor'
	config.buff_abbrev["Illusion of Life"] = 'iol'

	config.condition_ids = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
	config.auras_ids = {5677: 'Fire', 5577: 'Shocking', 5579: 'Frost', 5684: 'Magnetic'}

	config.charts = config_input.charts
	config.include_comp_and_review = config_input.include_comp_and_review
			
	return config
	
		
# For all players considered to be top in stat in this fight, increase
# the number of fights they reached top by 1 (i.e. increase
# consistency_stats[stat]).
# Input:
# players = list of all players
# sortedList = list of player names+profession, stat_value sorted by stat value in this fight
# config = configuration to use
# stat = stat that is considered
def increase_top_x_reached(players, sortedList, config, stat):
	valid_values = 0
	# filter out com for dist to tag
	if stat == 'dist':
		# different for dist
		first_valid = True
		i = 0
		last_val = 0
		while i < len(sortedList) and (valid_values < (len(sortedList)*config.num_players_considered_top)+1 or sortedList[i][1] == last_val):
			# sometimes dist is -1, filter these out
			if sortedList[i][1] >= 0:
				# first valid dist is the comm, don't consider
				if first_valid:
					first_valid  = False
				else:
					players[sortedList[i][0]].consistency_stats[stat] += 1
					valid_values += 1
			last_val = sortedList[i][1]
			i += 1
		return

	# total value doesn't need to be > 0 for deaths
	elif stat == 'deaths':
		i = 0
		last_val = 0
		while i < len(sortedList) and (valid_values < (len(sortedList)*config.num_players_considered_top) or sortedList[i][1] == last_val):
			if sortedList[i][1] < 0:
				i += 1
				continue
			if sortedList[i][1] == 0:
				players[sortedList[i][0]].consistency_stats[stat] += 1
				last_val = sortedList[i][1]
			i += 1
			valid_values += 1
		return
	
	
	# increase top stats reached for the first num_players_considered_top players
	i = 0
	last_val = 0
	while i < len(sortedList) and (valid_values < (len(sortedList)*config.num_players_considered_top) or sortedList[i][1] == last_val) and players[sortedList[i][0]].total_stats[stat] > 0:
		if sortedList[i][1] < 0 or (sortedList[i][1] == 0 and stat != 'dmg_taken'):
			i += 1
			continue
		players[sortedList[i][0]].consistency_stats[stat] += 1
		last_val = sortedList[i][1]
		i += 1
		valid_values += 1
	return



		
# sort the list of players by total value in stat
# Input:
# players = list of all Players
# stat = stat that is considered
# fight_num = number of the fight that is considered
# Output:
# list of player index and total stat value, sorted by total stat value
def sort_players_by_value_in_fight(players, stat, fight_num):
	decorated = [(player.stats_per_fight[fight_num][stat], i, player) for i, player in enumerate(players)]
	if stat == 'dist' or stat == 'dmg_taken' or stat == 'deaths':
		decorated.sort()
	else:
		decorated.sort(reverse=True)
	sorted_by_value = [(i, value) for value, i, player in decorated]
	return sorted_by_value



# sort the list of players by total value in stat
# Input:
# players = list of all Players
# stat = stat that is considered
# Output:
# list of player index and total stat value, sorted by total stat value
def sort_players_by_total(players, stat):
	decorated = [(player.total_stats[stat], i, player) for i, player in enumerate(players)]
	if stat == 'dist' or stat == 'dmg_taken' or stat == 'deaths':
		decorated.sort()
	else:
		decorated.sort(reverse=True)
	sorted_by_total = [(i, total) for total, i, player in decorated]
	return sorted_by_total



# sort the list of players by consistency value in stat
# Input:
# players = list of all Players
# stat = stat that is considered
# Output:
# list of player index and consistency stat value, sorted by consistency stat value (how often was top x reached)
def sort_players_by_consistency(players, stat):
	decorated = [(player.consistency_stats[stat], player.total_stats[stat], i, player) for i, player in enumerate(players)]
	decorated.sort(reverse=True)    
	sorted_by_consistency = [(i, consistency) for consistency, total, i, player in decorated]
	return sorted_by_consistency


# sort the list of players by percentage value in stat
# Input:
# players = list of all Players
# stat = stat that is considered
# Output:
# list of player index and percentage stat value, sorted by percentage stat value (how often was top x reached / number of fights attended)
def sort_players_by_percentage(players, stat):
	decorated = [(player.portion_top_stats[stat], player.consistency_stats[stat], player.total_stats[stat], i, player) for i, player in enumerate(players)]                
	decorated.sort(reverse=True)    
	sorted_by_percentage = [(i, percentage) for percentage, consistency, total, i, player in decorated]
	return sorted_by_percentage


# sort the list of players by average value in stat
# Input:
# players = list of all Players
# stat = stat that is considered
# Output:
# list of player index and average stat value, sorted by average stat value ( total stat value / duration of fights attended)
def sort_players_by_average(players, stat):
	decorated = [(player.average_stats[stat], player.consistency_stats[stat], player.total_stats[stat], i, player) for i, player in enumerate(players)]                
	if stat == 'dist' or stat == 'dmg_taken' or stat == 'deaths':
		decorated.sort()
	else:
		decorated.sort(reverse=True)    
	sorted_by_average = [(i, average) for average, consistency, total, i, player in decorated]
	return sorted_by_average



# Input:
# players = list of Players
# config = the configuration being used to determine top players
# stat = which stat are we considering
# total_or_consistent_or_average = enum StatType, either StatType.TOTAL, StatType.CONSISTENT or StatType.AVERAGE, we are getting the players with top total values, top consistency values, or top average values.
# Output:
# list of player indices getting a consistency / total / average award
def get_top_players(players, config, stat, total_or_consistent_or_average):
	percentage = 0.
	sorted_index = []
	if total_or_consistent_or_average == StatType.TOTAL and stat != 'dmg':
		percentage = float(config.portion_of_top_for_total)
		sorted_index = sort_players_by_total(players, stat)
	elif total_or_consistent_or_average == StatType.TOTAL and stat == 'dmg':
		percentage = float(config.portion_of_topDamage_for_total)
		sorted_index = sort_players_by_total(players, stat)		
	elif total_or_consistent_or_average == StatType.CONSISTENT:
		percentage = float(config.portion_of_top_for_consistent)
		sorted_index = sort_players_by_consistency(players, stat)
	elif total_or_consistent_or_average == StatType.AVERAGE and stat != 'dmg':
		percentage = float(config.portion_of_top_for_average)
		sorted_index = sort_players_by_average(players, stat)        
	elif total_or_consistent_or_average == StatType.AVERAGE and stat == 'dmg':
		percentage = float(config.portion_of_topDamage_for_total)
		sorted_index = sort_players_by_average(players, stat)        
	else:
		print("ERROR: Called get_top_players for stats that are not total or consistent")
		return        
		
	if config.player_sorting_stat_type == 'average':
		top_value = players[sorted_index[0][0]].average_stats[stat]
	else:
		top_value = players[sorted_index[0][0]].total_stats[stat]	
	top_players = list()

	i = 0
	last_value = 0
	while i < len(sorted_index):
		new_value = sorted_index[i][1] # value by which was sorted, i.e. total or consistency
		# index must be lower than number of output desired OR list entry has same value as previous entry, i.e. double place
		if i >= config.num_players_listed[stat] and new_value != last_value:
			break
		last_value = new_value

		if config.player_sorting_stat_type == 'average':
			stat_value = players[sorted_index[i][0]].average_stats[stat]
		else:
			stat_value = players[sorted_index[i][0]].total_stats[stat]
		# if stat isn't distance or dmg taken, total value must be at least percentage % of top value
		attendance_percentage = players[sorted_index[i][0]].attendance_percentage
		if (stat_value >= top_value*percentage and attendance_percentage > config.min_attendance_percentage_for_top) or (stat in ["dist", "dmg_taken"] and attendance_percentage > config.min_attendance_percentage_for_average):
			top_players.append(sorted_index[i][0])

		i += 1

	return top_players
			


# Input:
# players = list of Players
# config = the configuration being used to determine top players
# stat = which stat are we considering
# late_or_swapping = which type of stat. can be StatType.PERCENTAGE, StatType.LATE_PERCENTAGE or StatType.SWAPPED_PERCENTAGE
# num_used_fights = number of fights considered for computing top stats
# top_consistent_players = list of top consistent player indices
# top_total_players = list of top total player indices
# top_percentage_players = list of top percentage player indices
# top_late_players = list of player indices with late but great awards
# Output:
# list of player indices getting a percentage award, value with which the percentage stat was compared
def get_top_percentage_players(players, config, stat, late_or_swapping, num_used_fights, top_consistent_players, top_total_players, top_percentage_players, top_late_players):    
	sorted_index = sort_players_by_percentage(players, stat)
	top_percentage = players[sorted_index[0][0]].portion_top_stats[stat]
	
	comparison_value = 0
	min_attendance = 0
	if late_or_swapping == StatType.LATE_PERCENTAGE:
		comparison_value = top_percentage * config.portion_of_top_for_late
		min_attendance = config.min_attendance_portion_for_late * num_used_fights
	elif late_or_swapping == StatType.SWAPPED_PERCENTAGE:
		comparison_value = top_percentage * config.portion_of_top_for_buildswap
		min_attendance = config.min_attendance_portion_for_buildswap * num_used_fights
	elif late_or_swapping == StatType.PERCENTAGE:
		comparison_value = top_percentage * config.portion_of_top_for_percentage
		min_attendance = config.min_attendance_portion_for_percentage * num_used_fights
	else:
		print("ERROR: Called get_top_percentage_players for stats that are not percentage, late_percentage or swapped_percentage")
		return

	top_players = list()

	last_value = 0
	for (ind, percent) in sorted_index:
		# player wasn't there for enough fights
		if players[ind].num_fights_present < min_attendance:
			continue
		# player was there for all fights -> not late or swapping
		if late_or_swapping != StatType.PERCENTAGE and players[ind].num_fights_present == num_used_fights:
			continue
		# player got a different award already -> not late or swapping
		if late_or_swapping != StatType.PERCENTAGE and (ind in top_consistent_players or ind in top_total_players or ind in top_percentage_players or ind in top_late_players):
			continue
		# stat type swapping, but player didn't swap build
		if late_or_swapping == StatType.SWAPPED_PERCENTAGE and not players[ind].swapped_build:
			continue
		# index must be lower than number of output desired OR list entry has same value as previous entry, i.e. double place
		if len(top_players) >= config.num_players_listed[stat] and percent != last_value:
			break
		last_value = percent

		if percent >= comparison_value:
			top_players.append(ind)

	return top_players, comparison_value
 


# get the professions of all players indicated by the indices. Additionally, get the length of the longest profession name.
# Input:
# players = list of all players
# indices = list of relevant indices
# config = config to use for top stats computation/printing
# Output:
# list of profession strings, maximum profession length
def get_professions_and_length(players, indices, config):
	profession_strings = list()
	profession_length = 0
	for i in indices:
		player = players[i]
		professions_str = config.profession_abbreviations[player.profession]
		profession_strings.append(professions_str)
		if len(professions_str) > profession_length:
			profession_length = len(professions_str)
	return profession_strings, profession_length



# Get and write the top x people who achieved top y in stat most often.
# Input:
# players = list of Players
# config = the configuration being used to determine the top consistent players
# num_used_fights = the number of fights that are being used in stat computation
# stat = which stat are we considering
# output_file = the file to write the output to
# Output:
# list of player indices that got a top consistency award
def get_and_write_sorted_top_consistent(players, config, num_used_fights, stat, output_file):
	top_consistent_players = get_top_players(players, config, stat, StatType.CONSISTENT)
	write_sorted_top_consistent_or_avg(players, top_consistent_players, config, num_used_fights, stat, StatType.CONSISTENT, output_file)
	return top_consistent_players



# Get and write the people who achieved top x average in stat
# Input:
# players = list of Players
# config = the configuration being used to determine the top consistent players
# num_used_fights = the number of fights that are being used in stat computation
# stat = which stat are we considering
# output_file = the file to write the output to
# Output:
# list of player indices that got a top consistency award
def get_and_write_sorted_average(players, config, num_used_fights, stat, output_file):
	top_average_players = get_top_players(players, config, stat, StatType.AVERAGE)
	write_sorted_top_consistent_or_avg(players, top_average_players, config, num_used_fights, stat, StatType.AVERAGE, output_file)
	return top_average_players



#JEL - Modified for TW5 Output
# Write the top x people who achieved top y in stat most often.
# Input:
# players = list of Players
# top_consistent_players = list of Player indices considered top consistent players
# config = the configuration being used to determine the top consistent players
# num_used_fights = the number of fights that are being used in stat computation
# stat = which stat are we considering
# output_file = the file to write the output to
# Output:
# list of player indices that got a top consistency award
def write_sorted_top_consistent_or_avg(players, top_consistent_players, config, num_used_fights, stat, consistent_or_avg, output_file):
	#max_name_length = max([len(players[i].name) for i in top_consistent_players])
	profession_strings, profession_length = get_professions_and_length(players, top_consistent_players, config)

	if consistent_or_avg == StatType.CONSISTENT:
		if stat == "distJEL":
			print_string = "\n\n*Top "+str(config.num_players_considered_top*100)+"% "+config.stat_names[stat]+" consistency awards"
		else:
			print_string = "\n\n*Top "+config.stat_names[stat]+" consistency awards (Max. "+str(config.num_players_listed[stat])+" places, min. "+str(round(config.portion_of_top_for_consistent*100.))+"% of most consistent)"
			myprint(output_file, print_string)
			print_string = "*Most times placed in the top "+str(config.num_players_considered_top*100)+"%."
			myprint(output_file, print_string)
			print_string =  "*Attendance = number of fights a player was present out of "+str(num_used_fights)+" total fights."
			myprint(output_file, print_string)
	elif consistent_or_avg == StatType.AVERAGE:
		if stat == "distJEL":
			print_string = "*Top average "+str(config.num_players_considered_top*100)+"% "+config.stat_names[stat]+" awards"
		else:
			print_string = "*Top average "+config.stat_names[stat]+" awards (Max. "+str(config.num_players_listed[stat])+" places)"
			myprint(output_file, print_string)
			#print_string = "*''FightTime avg'': Total/total duration of fights | ''CombatTime avg''= Total/total time player alive during fights"
			#myprint(output_file, print_string)			
			print_string = "*Attendance = number of fights a player was present out of "+str(num_used_fights)+" total fights."
			myprint(output_file, print_string)
	print_string = "\n"    
	myprint(output_file, print_string)


	# print table header
	print_string = "|thead-dark table-hover sortable|k"    
	myprint(output_file, print_string)
	print_string = "|!Place |!Name |!Class | !Attendance| !Times Top|"
	if stat == 'dmg_taken':
		print_string += " !Total| !Average|"		
	print_string += "h"
	myprint(output_file, print_string)    

	
	place = 0
	last_val = 0
	# print table
	for i in range(len(top_consistent_players)):
		player = players[top_consistent_players[i]]
		#if player.duration_in_combat > 0:
		#	combat_Time = int(player.duration_in_combat)
		#else:
		#	combat_Time = int(player.duration_fights_present)
		if player.consistency_stats[stat] != last_val:
			place += 1
		nameWithTooltip = '<span data-tooltip="'+player.account+'">'+player.name+'</span>'
		print_string = "|"+str(place)+". |"+nameWithTooltip+" | {{"+profession_strings[i]+"}} | "+str(player.num_fights_present)+" | "+my_value(round(player.consistency_stats[stat]))+" |"
		if stat == 'dmg_taken':
			print_string += " "+my_value(round(player.total_stats[stat],1))+"| "+my_value(round(player.average_stats[stat]))+"|"

		myprint(output_file, print_string)
		last_val = player.consistency_stats[stat]
	myprint(output_file, "\n")
		
				
# Write out accounts that played support classes
def write_support_players(players, top_players, stat, output_file):
	for i in range(len(top_players)):
		player = players[top_players[i]]
		if Guild_Data:
			guildStatus = findMember(members, player.account)
		else:
			guildStatus = ""
		if stat == 'rips' and (player.profession == 'Chronomancer' or player.profession == 'Spellbreaker'):
			print_string = "|"+player.account+" |"+player.name+" |"+player.profession+" | "+str(player.num_fights_present)+"| "+str(player.duration_fights_present)+"| "+stat+" |"+guildStatus+" |"
			myprint(output_file, print_string)
		if stat == 'cleanses' and (player.profession == 'Scrapper' or player.profession == 'Tempest' or player.profession == 'Druid'):
			print_string = "|"+player.account+" |"+player.name+" |"+player.profession+" | "+str(player.num_fights_present)+"| "+str(player.duration_fights_present)+"| "+stat+" |"+guildStatus+" |"
			myprint(output_file, print_string)
		if stat == 'stability' and (player.profession == 'Firebrand'):
			print_string = "|"+player.account+" |"+player.name+" |"+player.profession+" | "+str(player.num_fights_present)+"| "+str(player.duration_fights_present)+"| "+stat+" |"+guildStatus+" |"
			myprint(output_file, print_string)
		if stat == 'heal' and (player.profession == 'Vindicator'):
			print_string = "|"+player.account+" |"+player.name+" |"+player.profession+" | "+str(player.num_fights_present)+"| "+str(player.duration_fights_present)+"| "+stat+" |"+guildStatus+" |"
			myprint(output_file, print_string)
# Write the top x people who achieved top total stat.
# Input:
# players = list of Players
# top_players = list of indices in players that are considered as top
# stat = which stat are we considering
# xls_output_filename = where to write to
def write_stats_xls(players, top_players, stat, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet(stat)
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Account")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Attendance (number of fights)")
	sheet1.write(0, 5, "Attendance (duration fights)")
	sheet1.write(0, 6, "Times Top")
	sheet1.write(0, 7, "Percentage Top")
	sheet1.write(0, 8, "Total "+stat)
	if stat == 'deaths':
		sheet1.write(0, 9, "Average "+stat+" per min")
	else:
		sheet1.write(0, 9, "Average "+stat+" per s")        

	for i in range(len(top_players)):
		player = players[top_players[i]]
		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, player.account)
		sheet1.write(i+1, 2, player.name)
		sheet1.write(i+1, 3, player.profession)
		sheet1.write(i+1, 4, player.num_fights_present)
		sheet1.write(i+1, 5, player.duration_fights_present)
		sheet1.write(i+1, 6, player.consistency_stats[stat])        
		sheet1.write(i+1, 7, round(player.portion_top_stats[stat]*100, 4))
		sheet1.write(i+1, 8, round(player.total_stats[stat], 4))
		sheet1.write(i+1, 9, player.average_stats[stat])

	wb.save(xls_output_filename)

def write_control_effects_out_xls(sorted_squadControl, stat, players, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet(stat+"- Out")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Place")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Total "+stat+" Outbound")
	sheet1.write(0, 5, "Average "+stat+" Outbound")	
	
	i = 0

	for name in sorted_squadControl:
		prof = "Not Found"
		fightTime = 999999
		
		for nameIndex in players:
			if nameIndex.name == name:
				prof = nameIndex.profession
				fightTime = nameIndex.duration_fights_present
		if i < 25:
			sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet1.write(i+1, 1, i+1)
			sheet1.write(i+1, 2, name)
			sheet1.write(i+1, 3, prof)
			sheet1.write(i+1, 4, round(sorted_squadControl[name], 1))
			sheet1.write(i+1, 5, round(sorted_squadControl[name]/fightTime, 4))
			i=i+1
	wb.save(xls_output_filename)

def write_control_effects_in_xls(sorted_enemyControl, stat, players, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet(stat+"- In")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Place")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Total "+stat+" Inbound")
	
	i = 0

	for name in sorted_enemyControl:
		prof = "Not Found"
		fightTime = 999999
		
		for nameIndex in players:
			if nameIndex.name == name:
				prof = nameIndex.profession
				fightTime = nameIndex.duration_fights_present
		if i < 25:
			sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet1.write(i+1, 1, i+1)
			sheet1.write(i+1, 2, name)
			sheet1.write(i+1, 3, prof)
			sheet1.write(i+1, 4, round(sorted_enemyControl[name], 4))
			sheet1.write(i+1, 5, round(sorted_enemyControl[name]/fightTime, 4))
			i=i+1
	wb.save(xls_output_filename)

def write_Attendance_xls(Attendance, xls_output_filename):

	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet("Attendance")

	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Account")
	sheet1.write(0, 2, "Attendance (number of fights)")
	sheet1.write(0, 3, "Attendance (duration fights)")	
	sheet1.write(0, 4, "Guild Status")	
	
	i=0
	
	for account in Attendance:
		Acct_Fights = Attendance[account]['fights']
		Acct_Duration = Attendance[account]['duration']
		Acct_Guild_Status = Attendance[account]['guildStatus']
		
		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, account)
		sheet1.write(i+1, 2, Acct_Fights)
		sheet1.write(i+1, 3, Acct_Duration)
		sheet1.write(i+1, 4, Acct_Guild_Status)
		i=i+1
	wb.save(xls_output_filename)

def write_Death_OnTag_xls(Death_OnTag, uptime_Table, players, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet("Death_OnTag")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Name")
	sheet1.write(0, 2, "Profession")
	sheet1.write(0, 3, "Attendance")
	sheet1.write(0, 4, "On_Tag")
	sheet1.write(0, 5, "Off_Tag")
	sheet1.write(0, 6, "After_Tag_Death")
	sheet1.write(0, 7, "Run_Back")
	sheet1.write(0, 8, "Total")
	sheet1.write(0, 9, "Off Tag Ranges")

		
	i = 0

	for prof_name in Death_OnTag:
		fightTime = uptime_Table[prof_name]['duration']
		if Death_OnTag[prof_name]['Off_Tag']:
			converted_list = [str(round(element)) for element in Death_OnTag[prof_name]['Ranges']]
			Ranges_string = ",".join(converted_list)
		else:
			Ranges_string = " "

		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, Death_OnTag[prof_name]['name'])
		sheet1.write(i+1, 2, Death_OnTag[prof_name]['profession'])
		sheet1.write(i+1, 3, fightTime)
		sheet1.write(i+1, 4, Death_OnTag[prof_name]['On_Tag'])
		sheet1.write(i+1, 5, Death_OnTag[prof_name]['Off_Tag'])
		sheet1.write(i+1, 6, Death_OnTag[prof_name]['After_Tag_Death'])
		sheet1.write(i+1, 7, Death_OnTag[prof_name]['Run_Back'])
		sheet1.write(i+1, 8, Death_OnTag[prof_name]['Total'])
		sheet1.write(i+1, 9, Ranges_string)
		i=i+1
	wb.save(xls_output_filename)

def write_DPSStats_xls(DPSStats, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet("DPSStats")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Account")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Role")
	sheet1.write(0, 5, "Attendance")
	sheet1.write(0, 6, "Combat Time")
	sheet1.write(0, 7, "Damage")
	sheet1.write(0, 8, "Squad Damage")
	sheet1.write(0, 9, "Downs")
	sheet1.write(0, 10, "Kills")
	sheet1.write(0, 11, "Coordination Damage")
	sheet1.write(0, 12, "Carrion Damage")
	sheet1.write(0, 13, "Squad Carrion Damage")
	for j in range(1, 21):
		sheet1.write(0, 13 + j, 'Chunk Damage (' + str(j) + ')')
	for j in range(1, 21):
		sheet1.write(0, 33 + j, 'Squad Chunk Damage (' + str(j) + ')')
		
	i = 0

	for name in DPSStats:
		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, DPSStats[name]['account'])
		sheet1.write(i+1, 2, DPSStats[name]['name'])
		sheet1.write(i+1, 3, DPSStats[name]['profession'])
		sheet1.write(i+1, 4, DPSStats[name]['role'])
		sheet1.write(i+1, 5, DPSStats[name]['duration'])
		sheet1.write(i+1, 6, DPSStats[name]['combatTime'])
		sheet1.write(i+1, 7, DPSStats[name]['Damage_Total'])
		sheet1.write(i+1, 8, DPSStats[name]['Squad_Damage_Total'])
		sheet1.write(i+1, 9, DPSStats[name]['Downs'])
		sheet1.write(i+1, 10, DPSStats[name]['Kills'])
		sheet1.write(i+1, 11, DPSStats[name]['Coordination_Damage'])
		sheet1.write(i+1, 12, DPSStats[name]['Carrion_Damage'])
		sheet1.write(i+1, 13, DPSStats[name]['Carrion_Damage_Total'])
		for j in range(1, 21):
			sheet1.write(i+1, 13 + j, DPSStats[name]['Chunk_Damage'][j])
		for j in range(1, 21):
			sheet1.write(i+1, 33 + j, DPSStats[name]['Chunk_Damage_Total'][j])
		i=i+1
	
	# Add BurstDPS sheet
	sheet2 = wb.add_sheet("Burst Damage")
	
	sheet2.write(0, 0, "Date")
	sheet2.write(0, 1, "Account")
	sheet2.write(0, 2, "Name")
	sheet2.write(0, 3, "Profession")
	sheet2.write(0, 4, "Role")
	for j in range(1, 21):
		sheet2.write(0, 4 + j, 'Burst Damage (' + str(j) + ')')
		
	i = 0

	for name in DPSStats:
		sheet2.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet2.write(i+1, 1, DPSStats[name]['account'])
		sheet2.write(i+1, 2, DPSStats[name]['name'])
		sheet2.write(i+1, 3, DPSStats[name]['profession'])
		sheet2.write(i+1, 4, DPSStats[name]['role'])
		for j in range(1, 21):
			sheet2.write(i+1, 4 + j, DPSStats[name]['Burst_Damage'][j])
		i=i+1
	
	# Add Ch5CaBurstDPS sheet
	sheet3 = wb.add_sheet("Ch5Ca Burst Damage")
	
	sheet3.write(0, 0, "Date")
	sheet3.write(0, 1, "Account")
	sheet3.write(0, 2, "Name")
	sheet3.write(0, 3, "Profession")
	sheet3.write(0, 4, "Role")
	for j in range(1, 21):
		sheet3.write(0, 4 + j, 'Ch5Ca Burst Damage (' + str(j) + ')')
		
	i = 0

	for name in DPSStats:
		sheet3.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet3.write(i+1, 1, DPSStats[name]['account'])
		sheet3.write(i+1, 2, DPSStats[name]['name'])
		sheet3.write(i+1, 3, DPSStats[name]['profession'])
		sheet3.write(i+1, 4, DPSStats[name]['role'])
		for j in range(1, 21):
			sheet3.write(i+1, 4 + j, DPSStats[name]['Ch5Ca_Burst_Damage'][j])
		i=i+1

	wb.save(xls_output_filename)

def write_squad_offensive_xls(squad_offensive, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet("Squad_Offensive")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Name")
	sheet1.write(0, 2, "Profession")
	sheet1.write(0, 3, "Critical %")
	sheet1.write(0, 4, "Flanking %")
	sheet1.write(0, 5, "Glancing %")
	sheet1.write(0, 6, "Moving %")
	sheet1.write(0, 7, "Blind #")
	sheet1.write(0, 8, "Interupt #")
	sheet1.write(0, 9, "Invulnerable #")
	sheet1.write(0, 10, "Evaded #")
	sheet1.write(0, 11, "Blocked #")
	sheet1.write(0, 12, "Critable Direct Damage Count'")
	sheet1.write(0, 13, "Connected Direct Damage Count")
	sheet1.write(0, 14, "Total Damage Count")
		
	i = 0

	for squadDps_prof_name in squad_offensive:

		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, squad_offensive[squadDps_prof_name]['name'])
		sheet1.write(i+1, 2, squad_offensive[squadDps_prof_name]['prof'])
		if squad_offensive[squadDps_prof_name]['stats']['criticalRate']:
			sheet1.write(i+1, 3, round((squad_offensive[squadDps_prof_name]['stats']['criticalRate']/squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount']), 4))
		else:
			sheet1.write(i+1, 3, 0.0000)
		if squad_offensive[squadDps_prof_name]['stats']['flankingRate']:
			sheet1.write(i+1, 4, round((squad_offensive[squadDps_prof_name]['stats']['flankingRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount']), 4))
		else:
			sheet1.write(i+1, 4, 0.0000)
		if squad_offensive[squadDps_prof_name]['stats']['glanceRate']:
			sheet1.write(i+1, 5, round((squad_offensive[squadDps_prof_name]['stats']['glanceRate']/squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount']), 4))
		else:
			sheet1.write(i+1, 5, 0.0000)
		if squad_offensive[squadDps_prof_name]['stats']['againstMovingRate']:
			sheet1.write(i+1, 6, round((squad_offensive[squadDps_prof_name]['stats']['againstMovingRate']/squad_offensive[squadDps_prof_name]['stats']['totalDamageCount']), 4))
		else:
			sheet1.write(i+1, 6, 0.0000)
		if squad_offensive[squadDps_prof_name]['stats']['missed']:
			sheet1.write(i+1, 7, squad_offensive[squadDps_prof_name]['stats']['missed'])
		else:
			sheet1.write(i+1, 7, 0)
		if squad_offensive[squadDps_prof_name]['stats']['interrupts']:
			sheet1.write(i+1, 8, squad_offensive[squadDps_prof_name]['stats']['interrupts'])
		else:
			sheet1.write(i+1, 8, 0)
		if squad_offensive[squadDps_prof_name]['stats']['invulned']:
			sheet1.write(i+1, 9, squad_offensive[squadDps_prof_name]['stats']['invulned'])
		else:
			sheet1.write(i+1, 9, 0)
		if squad_offensive[squadDps_prof_name]['stats']['evaded']:
			sheet1.write(i+1, 10, squad_offensive[squadDps_prof_name]['stats']['evaded'])
		else:
			sheet1.write(i+1, 10, 0)
		if squad_offensive[squadDps_prof_name]['stats']['blocked']:
			sheet1.write(i+1, 11, squad_offensive[squadDps_prof_name]['stats']['blocked'])
		else:
			sheet1.write(i+1, 11, 0)
		sheet1.write(i+1, 12, squad_offensive[squadDps_prof_name]['stats']['critableDirectDamageCount'])
		sheet1.write(i+1, 13, squad_offensive[squadDps_prof_name]['stats']['connectedDirectDamageCount'])
		sheet1.write(i+1, 14, squad_offensive[squadDps_prof_name]['stats']['totalDamageCount'])
		i=i+1
	wb.save(xls_output_filename)

def write_auras_out_xls(sorted_auras_TableOut, stat, players, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet(stat+" Aura-Out")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Place")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Total "+stat+" Aura-Out")
	
	i = 0

	for name in sorted_auras_TableOut:
		prof = "Not Found"
		fightTime = 999999
		for nameIndex in players:
			if nameIndex.name == name:
				prof = nameIndex.profession
				fightTime = nameIndex.duration_fights_present
		if i < 25:
			sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet1.write(i+1, 1, i+1)
			sheet1.write(i+1, 2, name)
			sheet1.write(i+1, 3, prof)
			sheet1.write(i+1, 4, round(sorted_auras_TableOut[name], 4))
			sheet1.write(i+1, 5, round(sorted_auras_TableOut[name]/fightTime, 4))
			i=i+1
	wb.save(xls_output_filename)

def write_buff_uptimes_in_xls(uptime_Table, players, uptime_Order, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	sheet1 = wb.add_sheet("Buff Uptimes")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Name")
	sheet1.write(0, 2, "Profession")
	sheet1.write(0, 3, "Attendance")
	sheet1.write(0, 4, "Stability")
	sheet1.write(0, 5, "Protection")
	sheet1.write(0, 6, "Aegis")
	sheet1.write(0, 7, "Might")
	sheet1.write(0, 8, "Fury")
	sheet1.write(0, 9, "Resistance")
	sheet1.write(0, 10, "Resolution")
	sheet1.write(0, 11, "Quickness")
	sheet1.write(0, 12, "Swiftness")
	sheet1.write(0, 13, "Alacrity")
	sheet1.write(0, 14, "Vigor")
	sheet1.write(0, 15, "Regeneration")
	
	i = 0
	
	for prof_name in uptime_Table:
		fightTime = uptime_Table[prof_name]['duration']
		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, uptime_Table[prof_name]['name'])
		sheet1.write(i+1, 2, uptime_Table[prof_name]['prof'])
		sheet1.write(i+1, 3, fightTime)

		x = 0
		for item in uptime_Order:
			if item in uptime_Table[prof_name]:
				buff_Time = uptime_Table[prof_name][item]
				try:
					sheet1.write(i+1, 4+x, round(((buff_Time / fightTime) * 100), 4))
				except:
					sheet1.write(i+1, 4+x, 0.00)
			else:
				sheet1.write(i+1, 4+x, 0.00)
			x=x+1
		i=i+1
	wb.save(xls_output_filename)

def write_stacking_buff_uptimes_in_xls(uptimeTable, xls_output_filename):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)

	# Add Might Stack sheet
	sheet1 = wb.add_sheet("Might Stack Uptime")
	
	sheet1.write(0, 0, "Date")
	sheet1.write(0, 1, "Account")
	sheet1.write(0, 2, "Name")
	sheet1.write(0, 3, "Profession")
	sheet1.write(0, 4, "Role")
	sheet1.write(0, 5, "Attendance")
	for j in range(0, 26):
		sheet1.write(0, 6 + j, 'Might (' + str(j) + ')')
		
	i = 0

	for name in uptimeTable:
		if 'might' not in uptimeTable[name]:
			continue

		sheet1.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet1.write(i+1, 1, uptimeTable[name]['account'])
		sheet1.write(i+1, 2, uptimeTable[name]['name'])
		sheet1.write(i+1, 3, uptimeTable[name]['profession'])
		sheet1.write(i+1, 4, uptimeTable[name]['role'])
		sheet1.write(i+1, 5, uptimeTable[name]['duration_might'])
		for j in range(0, 26):
			sheet1.write(i+1, 6 + j, uptimeTable[name]['might'][j])
		i=i+1

	# Add Stability Stack sheet
	sheet2 = wb.add_sheet("Stability Stack Uptime")
	
	sheet2.write(0, 0, "Date")
	sheet2.write(0, 1, "Account")
	sheet2.write(0, 2, "Name")
	sheet2.write(0, 3, "Profession")
	sheet2.write(0, 4, "Role")
	sheet2.write(0, 5, "Attendance")
	for j in range(0, 26):
		sheet2.write(0, 6 + j, 'Stability (' + str(j) + ')')
		
	i = 0

	for name in uptimeTable:
		if 'stability' not in uptimeTable[name]:
			continue

		sheet2.write(i+1, 0, fileDate.strftime("%Y-%m-%d"))
		sheet2.write(i+1, 1, uptimeTable[name]['account'])
		sheet2.write(i+1, 2, uptimeTable[name]['name'])
		sheet2.write(i+1, 3, uptimeTable[name]['profession'])
		sheet2.write(i+1, 4, uptimeTable[name]['role'])
		sheet2.write(i+1, 5, uptimeTable[name]['duration_stability'])
		for j in range(0, 26):
			sheet2.write(i+1, 6 + j, uptimeTable[name]['stability'][j])
		i=i+1

	wb.save(xls_output_filename)

def write_support_xls(players, top_players, stat, xls_output_filename, supportCount):
	fileDate = datetime.datetime.now()
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	supportCount = supportCount
		
	try:
		wb.add_sheet('Support')
	except:
		pass

	sheet2 = wb.get_sheet('Support')

	sheet2.write(0, 0, "Date")
	sheet2.write(0, 1, "Account")
	sheet2.write(0, 2, "Name")
	sheet2.write(0, 3, "Profession")
	sheet2.write(0, 4, "Attendance (number of fights)")
	sheet2.write(0, 5, "Attendance (duration fights)")
	sheet2.write(0, 6, "Support Stat")

	for i in range(len(top_players)):
		player = players[top_players[i]]
		if stat == 'rips' and (player.profession == 'Chronomancer' or player.profession == 'Spellbreaker'):
			sheet2.write(supportCount+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet2.write(supportCount+1, 1, player.account)
			sheet2.write(supportCount+1, 2, player.name)
			sheet2.write(supportCount+1, 3, player.profession)
			sheet2.write(supportCount+1, 4, player.num_fights_present)
			sheet2.write(supportCount+1, 5, player.duration_fights_present)
			sheet2.write(supportCount+1, 6, stat)
			supportCount +=1

		if stat == 'cleanses' and (player.profession == 'Scrapper' or player.profession == 'Tempest' or player.profession == 'Druid'):
			sheet2.write(supportCount+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet2.write(supportCount+1, 1, player.account)
			sheet2.write(supportCount+1, 2, player.name)
			sheet2.write(supportCount+1, 3, player.profession)
			sheet2.write(supportCount+1, 4, player.num_fights_present)
			sheet2.write(supportCount+1, 5, player.duration_fights_present)
			sheet2.write(supportCount+1, 6, stat)
			supportCount +=1

		if stat == 'stability' and (player.profession == 'Firebrand'):
			sheet2.write(supportCount+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet2.write(supportCount+1, 1, player.account)
			sheet2.write(supportCount+1, 2, player.name)
			sheet2.write(supportCount+1, 3, player.profession)
			sheet2.write(supportCount+1, 4, player.num_fights_present)
			sheet2.write(supportCount+1, 5, player.duration_fights_present)
			sheet2.write(supportCount+1, 6, stat)
			supportCount +=1

		if stat == 'heal' and (player.profession == 'Vindicator'):
			sheet2.write(supportCount+1, 0, fileDate.strftime("%Y-%m-%d"))
			sheet2.write(supportCount+1, 1, player.account)
			sheet2.write(supportCount+1, 2, player.name)
			sheet2.write(supportCount+1, 3, player.profession)
			sheet2.write(supportCount+1, 4, player.num_fights_present)
			sheet2.write(supportCount+1, 5, player.duration_fights_present)
			sheet2.write(supportCount+1, 6, stat)
			supportCount +=1

	wb.save(xls_output_filename)
	return supportCount

# Get and write the top x people who achieved top total stat.
# Input:
# players = list of Players
# config = the configuration being used to determine topx consistent players
# total_fight_duration = the total duration of all fights
# stat = which stat are we considering
# output_file = where to write to
# Output:
# list of top total player indices
def get_and_write_sorted_total(players, config, total_fight_duration, stat, output_file):
	# get players that get an award and their professions
	top_total_players = get_top_players(players, config, stat, StatType.TOTAL)
	write_sorted_total(players, top_total_players, config, total_fight_duration, stat, output_file)
	return top_total_players

# Get and write the top x people who achieved top total stat.
# Input:
# players = list of Players
# config = the configuration being used to determine topx consistent players
# total_fight_duration = the total duration of all fights
# stat = which stat are we considering
# output_file = where to write to
# Output:
# list of top total player indices
def get_and_write_sorted_total_by_average(players, config, total_fight_duration, stat, output_file):
	# get players that get an award and their professions
	top_total_players = get_top_players(players, config, stat, StatType.AVERAGE)
	write_sorted_total(players, top_total_players, config, total_fight_duration, stat, output_file)
	return top_total_players



# Write the top x people who achieved top total stat.
# Input:
# players = list of Players
# top_total_players = list of Player indices considered top total players
# config = the configuration being used to determine topx consistent players
# total_fight_duration = the total duration of all fights
# stat = which stat are we considering
# output_file = where to write to
# Output:
# list of top total player indices
def write_sorted_total(players, top_total_players, config, total_fight_duration, stat, output_file):
	profession_strings, profession_length = get_professions_and_length(players, top_total_players, config)
	profession_length = max(profession_length, 5)
	if stat == 'dmg':
		print_string = "*Top overall "+config.stat_names[stat]+" awards (Max. "+str(config.num_players_listed[stat])+" places, min. "+str(round(config.portion_of_topDamage_for_total*100.))+"% of 1st place)"
	else:
		print_string = "*Top overall "+config.stat_names[stat]+" awards (Max. "+str(config.num_players_listed[stat])+" places, min. "+str(round(config.portion_of_top_for_total*100.))+"% of 1st place)"
	myprint(output_file, print_string)
	#print_string = "*''FightTime avg'': Total/total duration of fights | ''CombatTime avg''= Total/total time player alive during fights"
	#myprint(output_file, print_string)
	print_string = "*Attendance = total duration of fights attended out of "
	if total_fight_duration["h"] > 0:
		print_string += str(total_fight_duration["h"])+"h "
	print_string += str(total_fight_duration["m"])+"m "+str(total_fight_duration["s"])+"s."    
	myprint(output_file, print_string)
	print_string = "\n"
	myprint(output_file, print_string)

	#JEL - Adjust for TW5 table output
	#print_string = "|Place |Name |Class | Attendance| Total| "
	#    print_string += " Average|h"
	# print table header
	print_string = "|thead-dark table-hover sortable|k"
	myprint(output_file, print_string)
	print_string = "|!Place |!Name |!Class | !Attendance| !Total|"
	if stat in config.buff_ids:
		if stat == 'iol':
			print_string += " !FightTime Avg| !CombatTime Avg|"
		else:
			per_sec_name = stat[:1].upper() + "PS"
			print_string += f" !Squad {per_sec_name}| !Group {per_sec_name}| !Self {per_sec_name}|"
	elif stat == 'dmg':
		print_string += " !FightTime DPS| !CombatTime DPS|  !Damage/Enemy|   !Wt. DPS/Enemy|"
	elif stat == 'heal':
		print_string += " !Squad HPS| !Group HPS| !Self HPS|"
	elif stat == 'rips' or stat == 'rips-In':
		print_string += " !FightTime SPS| !CombatTime SPS|"
	elif stat == 'cleanses' or stat == 'cleanses-In':
		print_string += " !FightTime CPS| !CombatTime CPS|"
	elif stat == 'barrier':
		print_string += " !Squad BPS| !Group BPS| !Self BPS|"
	elif stat == 'downs':
		print_string += " !FightTime Downs/Min| !CombatTime Downs/Min|"
	elif stat == 'kills':
		print_string += " !FightTime Kills/Min| !CombatTime Kills/Min|"
	else:
		if stat in ['Pdmg', 'Cdmg', 'dmg_taken', 'barrierDamage', 'cleansesIn', 'ripsIn']:
			print_string += " !Per sec avg|"
		else:
			print_string += " !Per min avg|"
	print_string += "h"
	myprint(output_file, print_string)    

	place = 0
	last_val = -1
	# print table
	for i in range(len(top_total_players)):
		player = players[top_total_players[i]]
		if player.total_stats[stat] != last_val:
			place += 1

		fight_time_h = int(player.duration_fights_present/3600)
		fight_time_m = int((player.duration_fights_present - fight_time_h*3600)/60)
		fight_time_s = int(player.duration_fights_present - fight_time_h*3600 - fight_time_m*60)
		if player.duration_in_combat > 0:
			combat_Time = int(player.duration_in_combat)
		else:
			combat_Time = int(player.duration_fights_present)

		#JEL - Adjust for TW5 table output
		nameWithTooltip = '<span data-tooltip="'+player.account+'">'+player.name+'</span>'
		print_string = "|"+str(place)+". |"+nameWithTooltip+" | {{"+profession_strings[i]+"}} | "
		#print_string = f"{place:>2}"+f". {player.name:<{max_name_length}} "+f" {profession_strings[i]:<{profession_length}} "

		if fight_time_h > 0:
			print_string += f" {fight_time_h:>2}h {fight_time_m:>2}m {fight_time_s:>2}s |"
		else:
			print_string += f" {fight_time_m:>6}m {fight_time_s:>2}s |"
		if stat == 'cleanses' or stat == 'rips' or stat == 'cleanses-In' or stat == 'rips-In':
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			print_string += " "+"{:.2f}".format(round(player.average_stats[stat], 2))+"| "+"{:.2f}".format(round(player.total_stats[stat]/combat_Time, 2))+"|"
		elif stat == 'barrier' or stat == 'heal':
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			print_string += " "+"{:.2f}".format(round(player.average_stats[stat], 2))+"|"
			print_string += " "+"{:.2f}".format(round(player.total_stats_group[stat]/player.duration_fights_present, 2))+"|"
			print_string += " "+"{:.2f}".format(round(player.total_stats_self[stat]/player.duration_fights_present, 2))+"|"
		elif stat == 'dmg':
			weighted_DPS_Enemy=[]
			for (dmg, enemy, duration) in zip(player.wt_dps_damage, player.wt_dps_enemies, player.wt_dps_duration):
				if dmg != -1:
					weighted_DPS_Enemy.append(round((dmg/enemy/duration) * (duration / sum(player.wt_dps_duration)),4))
			wt_dps_enemy = sum(weighted_DPS_Enemy)
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			print_string += " "+my_value(round(player.average_stats[stat]))+"| "+my_value(round(player.total_stats[stat]/combat_Time))+"| "+my_value(round(player.total_stats[stat]/(player.num_enemies_present)))+"| "+my_value(round(wt_dps_enemy,4))+"|"
		elif stat == 'downs' or stat == 'kills':
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			print_string += " "+"{:.2f}".format(round(player.average_stats[stat]*60, 2))+"| "+"{:.2f}".format(round((player.total_stats[stat]/combat_Time)*60, 2))+"|"
		elif stat == 'iol':
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			print_string += " "+"{:.2f}".format(round(player.average_stats[stat], 2))+"| "+"{:.2f}".format(round((player.total_stats[stat]/combat_Time)*100, 2))+"|"
		elif stat in config.buffs_stacking_duration:
			if  player.duration_fights_present >0 and player.num_fights_present >0 and player.num_allies_supported >0:
				stat_Generated_Squad = (player.total_stats[stat]/((player.num_allies_supported - player.num_fights_present)/player.num_fights_present)/ player.duration_fights_present)*100
			else:
				stat_Generated_Squad = 0
			if  player.duration_fights_present >0 and player.num_fights_present >0 and player.num_allies_group_supported >0:
				if (player.num_allies_group_supported - player.num_fights_present) != 0:
					stat_Generated_Group = (player.total_stats_group[stat]/((player.num_allies_group_supported - player.num_fights_present)/player.num_fights_present)/ player.duration_fights_present)*100
				else:
					stat_Generated_Group = 0
			else:
				stat_Generated_Group = 0
			if  player.duration_fights_present >0:	
				stat_Generated_Self = (player.total_stats_self[stat]/player.duration_fights_present)*100
			else:
				stat_Generated_Self = 0

			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Squad, 4))+'% Squad Generation">'+my_value(round(player.total_stats[stat]))+"</span>|"
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Squad, 4))+'% Squad Generation">'+"{:.2f}".format(round(player.average_stats[stat], 2))+'</span>|'
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Group, 4))+'% Group Generation">'+"{:.2f}".format(round(player.total_stats_group[stat]/player.duration_fights_present, 2))+'</span>|'
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Self, 4))+'% Self Generation">'+"{:.2f}".format(round(player.total_stats_self[stat]/player.duration_fights_present, 2))+'</span>|'
		elif stat in config.buffs_stacking_intensity:
			if  player.duration_fights_present >0 and player.num_fights_present >0 and player.num_allies_supported >0:
				stat_Generated_Squad = (player.total_stats[stat]/((player.num_allies_supported - player.num_fights_present)/player.num_fights_present)/ player.duration_fights_present)
			else:
				stat_Generated_Squad = 0
			if  player.duration_fights_present >0 and player.num_fights_present >0 and player.num_allies_group_supported >0:
				if (player.num_allies_group_supported - player.num_fights_present) != 0:
					stat_Generated_Group = (player.total_stats_group[stat]/((player.num_allies_group_supported - player.num_fights_present)/player.num_fights_present)/ player.duration_fights_present)
			else:
				stat_Generated_Group = 0
			if  player.duration_fights_present >0:
				stat_Generated_Self = (player.total_stats_self[stat]/player.duration_fights_present)
			else:
				stat_Generated_Self = 0

			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Squad, 4))+' Squad Generation">'+my_value(round(player.total_stats[stat]))+"</span>|"
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Squad, 4))+' Squad Generation">'+"{:.2f}".format(round(player.average_stats[stat], 2))+'</span>|'
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Group, 4))+' Group Generation">'+"{:.2f}".format(round(player.total_stats_group[stat]/player.duration_fights_present, 2))+'</span>|'
			print_string += " "+'<span data-tooltip="'+my_value(round(stat_Generated_Self, 4))+' Self Generation">'+"{:.2f}".format(round(player.total_stats_self[stat]/player.duration_fights_present, 2))+'</span>|'
		else:
			print_string += " "+my_value(round(player.total_stats[stat]))+"|"
			if stat in ['Pdmg', 'Cdmg', 'dmg_taken', 'barrierDamage', 'cleansesIn', 'ripsIn', 'deaths']:
				print_string += " "+"{:.2f}".format(round(player.average_stats[stat], 2))+"|"
			else:
				print_string += " "+"{:.2f}".format(round(player.average_stats[stat] * 60.0, 2))+"|"
		myprint(output_file, print_string)
		last_val = player.total_stats[stat]
	myprint(output_file, "\n")
	
   

# Get and write the top x people who achieved top in stat with the highest percentage. This only considers fights where each player was present, i.e., a player who was in 4 fights and achieved a top spot in 2 of them gets 50%, as does a player who was only in 2 fights and achieved a top spot in 1 of them.
# Input:
# players = list of Players
# config = the configuration being used to determine topx consistent players
# num_used_fights = the number of fights that are being used in stat computation
# stat = which stat are we considering
# output_file = file to write to
# late_or_swapping = which type of stat. can be StatType.PERCENTAGE, StatType.LATE_PERCENTAGE or StatType.SWAPPED_PERCENTAGE
# top_consistent_players = list with indices of top consistent players
# top_total_players = list with indices of top total players
# top_percentage_players = list with indices of players with top percentage award
# top_late_players = list with indices of players who got a late but great award
# Output:
# list of players that got a top percentage award (or late but great or jack of all trades)
def get_and_write_sorted_top_percentage(players, config, num_used_fights, stat, output_file, late_or_swapping, top_consistent_players, top_total_players = list(), top_percentage_players = list(), top_late_players = list()):
	# get names that get on the list and their professions
	top_percentage_players, comparison_percentage = get_top_percentage_players(players, config, stat, late_or_swapping, num_used_fights, top_consistent_players, top_total_players, top_percentage_players, top_late_players)
	write_sorted_top_percentage(players, top_percentage_players, comparison_percentage, config, num_used_fights, stat, output_file, late_or_swapping, top_consistent_players, top_total_players, top_percentage_players, top_late_players)
	return top_percentage_players, comparison_percentage


# Write the top x people who achieved top in stat with the highest percentage. This only considers fights where each player was present, i.e., a player who was in 4 fights and achieved a top spot in 2 of them gets 50%, as does a player who was only in 2 fights and achieved a top spot in 1 of them.
# Input:
# players = list of Players
# top_players = list of Player indices considered top percentage players
# config = the configuration being used to determine topx consistent players
# num_used_fights = the number of fights that are being used in stat computation
# stat = which stat are we considering
# output_file = file to write to
# late_or_swapping = which type of stat. can be StatType.PERCENTAGE, StatType.LATE_PERCENTAGE or StatType.SWAPPED_PERCENTAGE
# top_consistent_players = list with indices of top consistent players
# top_total_players = list with indices of top total players
# top_percentage_players = list with indices of players with top percentage award
# top_late_players = list with indices of players who got a late but great award
# Output:
# list of players that got a top percentage award (or late but great or jack of all trades)
def write_sorted_top_percentage(players, top_players, comparison_percentage, config, num_used_fights, stat, output_file, late_or_swapping, top_consistent_players, top_total_players = list(), top_percentage_players = list(), top_late_players = list()):
	# get names that get on the list and their professions
	if len(top_players) <= 0:
		return top_players

	profession_strings, profession_length = get_professions_and_length(players, top_players, config)
	max_name_length = max([len(players[i].name) for i in top_players])
	profession_length = max(profession_length, 5)

	# print table header
	print_string = "`Top "+config.stat_names[stat]+" percentage (Minimum percentage = "+f"{comparison_percentage*100:.0f}%)"
	myprint(output_file, print_string)
	print_string = "`\n"     
	myprint(output_file, print_string)                

	# print table header
	print_string = "|thead-dark table-hover sortable|k"
	myprint(output_file, print_string)
	print_string = "|!Place |!Name |!Class | !Percentage | !Times Top | !Out of |"
	if stat != "dist":
		print_string += " !Total|h"
	else:
		print_string += "h"
	myprint(output_file, print_string)    

	# print stats for top players
	place = 0
	last_val = 0
	# print table
	for i in range(len(top_players)):
		player = players[top_players[i]]
		if player.portion_top_stats[stat] != last_val:
			place += 1

		percentage = int(player.portion_top_stats[stat]*100)
		nameWithTooltip = '<span data-tooltip="'+player.account+'">'+player.name+'</span>'
		#print_string = f"|{place:>2}"+f". |{player.name:<{max_name_length}} "+" | {{"+profession_strings[i]+"}} "+f"| {percentage:>10}% " +f" | {round(player.consistency_stats[stat]):>9} "+f" | {player.num_fights_present:>6} |"
		print_string = "|"+str(place)+". |"+nameWithTooltip+" | {{"+profession_strings[i]+"}} | "+str(percentage)+"%| "+str(round(player.consistency_stats[stat]))+" | "+ str(player.num_fights_present)+" |"
		if stat != "dist":
			print_string += f" {round(player.total_stats[stat]):>7} |"
		myprint(output_file, print_string)
		last_val = player.portion_top_stats[stat]
	myprint(output_file, "\n")


# get account, character name and profession from json object
def get_basic_player_data_from_json(player_json):
	account = player_json['account']
	name = player_json['name']
	profession = player_json['profession']
	return account, name, profession


def get_buff_ids_from_json(json_data, config):
	buffs = json_data['buffMap']
	for buff_id, buff in buffs.items():
		if buff['name'] in config.buff_abbrev:
			abbrev_name = config.buff_abbrev[buff['name']]
			config.buff_ids[abbrev_name] = buff_id[1:]
			if buff['stacking']:
				config.buffs_stacking_intensity.append(abbrev_name)
			else:
				config.buffs_stacking_duration.append(abbrev_name)
	#Quick fix for Buffs not found in the initial fight log buffMap
	BuffIdFix = { 'iol': 10346, 'superspeed': 5974,  'stealth': 13017,  'HiS': 10269,  'stability': 1122,  'protection': 717,  'aegis': 743,  'might': 740,  'fury': 725,  'resistance': 26980,  'resolution': 873,  'quickness': 1187,  'swiftness': 719,  'alacrity': 30328,  'vigor': 726,  'regeneration': 718, 'fireOut': 5677, 'shockingOut': 5577, 'frostOut': 5579, 'magneticOut':5684, 'lightOut': 25518, 'darkOut': 39978, 'chaosOut': 10332}
	for buff in BuffIdFix:
		if buff not in config.buff_ids:
			config.buff_ids[buff] = BuffIdFix[buff]
			if buff == 'might' or buff == 'stability':
				config.buffs_stacking_intensity.append(buff)
			else:
				config.buffs_stacking_duration.append(buff)


# Collect the top stats data.
# Input:
# args = cmd line arguments
# config = configuration to use for top stats computation
# log = log file to write to
# Output:
# list of Players with their stats
# list of all fights (also the skipped ones)
# was healing found in the logs?
def collect_stat_data(args, config, log, anonymize=False):
#    if args.filetype != "json" and args.filetype != "xml":
#        print("unsupported filetype "+args.filetype+". Please choose json or xml.")

	# healing only in logs if addon was installed
	found_healing = False # Todo what if some logs have healing and some don't
	found_barrier = False    

	players = []        # list of all player/profession combinations
	player_index = {}   # dictionary that matches each player/profession combo to its index in players list
	account_index = {}  # dictionary that matches each account name to a list of its indices in players list
	squad_comp = {}     # dictionary that contains count of professions by fight_num
	used_fights = 0
	fights = []
	first = True

	# iterating over all fights in directory
	files = listdir(args.input_directory)
	sorted_files = sorted(files)
	for filename in sorted_files:
		# skip files of incorrect filetype
		file_start, file_extension = os.path.splitext(filename)
		#if args.filetype not in file_extension or "top_stats" in file_start:
		if file_extension not in ['.json', '.gz'] or "top_stats" in file_start:
			continue

		print_string = "parsing "+filename
		print(print_string)
		file_path = "".join((args.input_directory,"/",filename))

		# load file
#        if args.filetype == "xml":
#            # create xml tree
#            xml_tree = ET.parse(file_path)
#            xml_root = xml_tree.getroot()
#            # get fight stats
#            fight, players_running_healing_addon = get_stats_from_fight_xml(xml_root, config, log)
#        else: # filetype == "json"
		if file_extension == '.gz':
			with gzip.open(file_path, mode="r") as f:
				json_data = json.loads(f.read().decode('utf-8'))
		else:
			json_datafile = open(file_path, encoding='utf-8')
			json_data = json.load(json_datafile)
		# get fight stats
		fight, players_running_healing_addon, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats = get_stats_from_fight_json(json_data, config, log)
			
		if first:
			first = False
			#if args.filetype == "json":
			get_buff_ids_from_json(json_data, config)
			#else:
			#    get_buff_ids_from_xml(xml_root, config)
					
		# add new entry for this fight in all players
		for player in players:
			player.stats_per_fight.append({key: value for key, value in config.empty_stats.items()})   

		fight_number = int(len(fights))
		# don't compute anything for skipped fights
		if fight.skipped:
			fights.append(fight)
			log.write("skipped "+filename)            
			continue

		party_member_counts = {}
		for player_data in json_data['players']:
			group = player_data['group']
			if group in party_member_counts:
				party_member_counts[group] += 1
			else:
				party_member_counts[group] = 1

		used_fights += 1
		#fight_number = used_fights-1
		squad_comp[fight_number]={}
		# get stats for each player
		#for player_data in (xml_root.iter('players') if args.filetype == "xml" else json_data['players']):
		for player_data in json_data['players']:
			create_new_player = False
			build_swapped = False

			#if args.filetype == "xml":
			#    account, name, profession = get_basic_player_data_from_xml(player_data)
			#else:
			account, name, profession = get_basic_player_data_from_json(player_data)                
			if profession not in squad_comp[fight_number]:
				squad_comp[fight_number][profession] = 1
			else:
				squad_comp[fight_number][profession] = squad_comp[fight_number][profession]+1

			# if this combination of charname + profession is not in the player index yet, create a new entry
			name_and_prof = name+" "+profession
			if name_and_prof not in player_index.keys():
				print("creating new player",name_and_prof)
				create_new_player = True

			# if this account is not in the account index yet, create a new entry
			if account not in account_index.keys():
				account_index[account] = [len(players)]
			elif name_and_prof not in player_index.keys():
				# if account does already exist, but name/prof combo does not, this player swapped build or character
				# -> note for all Player instances of this account
				for ind in range(len(account_index[account])):
					players[account_index[account][ind]].swapped_build = True
				account_index[account].append(len(players))
				build_swapped = True

			if create_new_player:
				player = Player(account, name, profession)
				player.initialize(config)
				player_index[name_and_prof] = len(players)
				# fill up fights where the player wasn't there yet with empty stats
				while len(player.stats_per_fight) <= fight_number:                
				#while len(player.stats_per_fight) <= used_fights:
					player.stats_per_fight.append({key: value for key, value in config.empty_stats.items()})                
				players.append(player)

			player = players[player_index[name_and_prof]]

			#if args.filetype == "xml":
			#    player.stats_per_fight[fight_number]['time_active'] = get_stat_from_player_xml(player_data, players_running_healing_addon, 'time_active', config)
			#else:
			player.stats_per_fight[fight_number]['time_active'] = get_stat_from_player_json(player_data, players_running_healing_addon, 'time_active', config)
			player.stats_per_fight[fight_number]['time_in_combat'] = get_stat_from_player_json(player_data, players_running_healing_addon, 'time_in_combat', config)
			player.stats_per_fight[fight_number]['group'] = get_stat_from_player_json(player_data, players_running_healing_addon, 'group', config)
			num_party_members = party_member_counts[player_data['group']]
			player.num_allies_group_supported += party_member_counts[player_data['group']]
			
			# get all stats that are supposed to be computed from the player data
			for stat in config.stats_to_compute:
				#if args.filetype == "xml":
				#    player.stats_per_fight[fight_number][stat] = get_stat_from_player_xml(player_data, players_running_healing_addon, stat, config)
				#else:
				player.stats_per_fight[fight_number][stat] = get_stat_from_player_json(player_data, players_running_healing_addon, stat, config)
					
				if stat == 'heal' and player.stats_per_fight[fight_number][stat] >= 0:
					found_healing = True
				elif stat == 'barrier' and player.stats_per_fight[fight_number][stat] >= 0:
					found_barrier = True                    
				elif stat == 'dist':
					player.stats_per_fight[fight_number][stat] = round(player.stats_per_fight[fight_number][stat])
				elif stat == 'dmg_taken':
					if player.stats_per_fight[fight_number]['time_in_combat'] == 0:
						player.stats_per_fight[fight_number]['time_in_combat'] = 1
					player.stats_per_fight[fight_number][stat] = player.stats_per_fight[fight_number][stat]/player.stats_per_fight[fight_number]['time_in_combat']

				#print(stat, name)
				# add stats of this fight and player to total stats of this fight and player
				if player.stats_per_fight[fight_number][stat] > 0:
					# buff are generation squad values, using total over time
					if stat in config.buffs_stacking_duration and stat != 'iol':
						#value is generated boon time on all squad players / fight duration / (players-1)" in percent, we want generated boon time on all squad players
						fight.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]/100.*fight.duration*(fight.allies - 1), 4)
						player.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]/100.*fight.duration*(fight.allies - 1), 4)

						group_gen = get_stat_from_player_json(player_data, players_running_healing_addon, stat, config, False, BuffGenerationType.GROUP)
						player.total_stats_group[stat] += round(group_gen/100.*fight.duration*(num_party_members - 1), 4)
						
						self_gen = get_stat_from_player_json(player_data, players_running_healing_addon, stat, config, False, BuffGenerationType.SELF)
						player.total_stats_self[stat] += round(self_gen/100.*fight.duration, 4)
					elif stat in config.buffs_stacking_intensity and stat != 'iol':
						#value is generated boon time on all squad players / fight duration / (players-1)", we want generated boon time on all squad players
						fight.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*fight.duration*(fight.allies - 1), 4)
						player.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*fight.duration*(fight.allies - 1), 4)
						
						group_gen = get_stat_from_player_json(player_data, players_running_healing_addon, stat, config, False, BuffGenerationType.GROUP)
						player.total_stats_group[stat] += round(group_gen*fight.duration*(num_party_members - 1), 4)
						
						self_gen = get_stat_from_player_json(player_data, players_running_healing_addon, stat, config, False, BuffGenerationType.SELF)
						player.total_stats_self[stat] += round(self_gen*fight.duration, 4)
					elif stat == 'dist':
						fight.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*fight.duration)
						player.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*fight.duration)
					elif stat == 'dmg_taken':
						fight.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*player.stats_per_fight[fight_number]['time_in_combat'])
						player.total_stats[stat] += round(player.stats_per_fight[fight_number][stat]*player.stats_per_fight[fight_number]['time_in_combat'])
					elif stat == 'heal':
						fight.total_stats[stat] += player.stats_per_fight[fight_number][stat]
						player.total_stats[stat] += player.stats_per_fight[fight_number][stat]

						if player_data['name'] in players_running_healing_addon and 'extHealingStats' in player_data:
							allied_healing_1s = player_data['extHealingStats']['alliedHealing1S']
							total_healing_group = 0
							for index in range(len(json_data['players'])):
								is_same_group = player_data['group'] == json_data['players'][index]['group']
								if is_same_group:
									total_healing_group += allied_healing_1s[index][0][-1]
									if player_data['name'] == json_data['players'][index]['name']:
										player.total_stats_self[stat] += allied_healing_1s[index][0][-1]
						player.total_stats_group[stat] += total_healing_group
					elif stat == 'barrier':
						fight.total_stats[stat] += player.stats_per_fight[fight_number][stat]
						player.total_stats[stat] += player.stats_per_fight[fight_number][stat]

						if player_data['name'] in players_running_healing_addon and 'extBarrierStats' in player_data:
							allied_barrier_1s = player_data['extBarrierStats']['alliedBarrier1S']
							total_barrier_group = 0
							for index in range(len(json_data['players'])):
								is_same_group = player_data['group'] == json_data['players'][index]['group']
								if is_same_group:
									total_barrier_group += allied_barrier_1s[index][0][-1]
									if player_data['name'] == json_data['players'][index]['name']:
										player.total_stats_self[stat] += allied_barrier_1s[index][0][-1]
						player.total_stats_group[stat] += total_barrier_group
					else:
						# all non-buff stats
						fight.total_stats[stat] += player.stats_per_fight[fight_number][stat]
						player.total_stats[stat] += player.stats_per_fight[fight_number][stat]
					
			if debug:
				print("\n")
				print(name)
				for stat in player.stats_per_fight[fight_number].keys():
					print(stat+": "+str(player.stats_per_fight[fight_number][stat]))
				print("\n\n")

			player.num_fights_present += 1
			player.num_enemies_present += fight.enemies
			player.num_allies_supported += (fight.allies)
			player.wt_dps_enemies.append(fight.enemies)
			player.wt_dps_duration.append(player.stats_per_fight[fight_number]['time_in_combat'])
			player.wt_dps_damage.append(player.stats_per_fight[fight_number]['dmg'])
			player.duration_fights_present += fight.duration
			player.duration_active += player.stats_per_fight[fight_number]['time_active']
			player.duration_in_combat += player.stats_per_fight[fight_number]['time_in_combat']
			player.swapped_build |= build_swapped
			player.stats_per_fight[fight_number]['fight_duration'] = fight.duration
			player.stats_per_fight[fight_number]['allies'] = fight.allies

		# create lists sorted according to stats
		sortedStats = {key: list() for key in config.stats_to_compute}
		for stat in config.stats_to_compute:
			sortedStats[stat] = sort_players_by_value_in_fight(players, stat, fight_number)

		if debug:
			for stat in config.stats_to_compute:
				print("sorted "+stat+": "+str(sortedStats[stat]))
		
		# increase number of times top x was achieved for top x players in each stat
		for stat in config.stats_to_compute:
			increase_top_x_reached(players, sortedStats[stat], config, stat)
			# round total_stats for this fight
			fight.total_stats[stat] = round(fight.total_stats[stat])

		fights.append(fight)

	if used_fights == 0:
		#print("ERROR: no valid fights with filetype "+args.filetype+" found in "+args.input_directory)
		print("ERROR: no valid fights with filetype json found in "+args.input_directory)
		exit(1)

	# compute percentage top stats and attendance percentage for each player    
	for player in players:
		player.attendance_percentage = round(player.num_fights_present / used_fights*100)
		# round total and portion top stats
		for stat in config.stats_to_compute:
			player.portion_top_stats[stat] = round(player.consistency_stats[stat]/player.num_fights_present, 4)
			player.total_stats[stat] = round(player.total_stats[stat], 4)
			if stat == 'dmg':
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present)
			elif stat == 'heal' or stat == 'barrier':
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present, 4)
			elif stat == 'dmg_taken':
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_in_combat)                
			elif stat == 'deaths':
				player.average_stats[stat] = round(player.total_stats[stat]/(player.duration_fights_present/60), 4)
			elif stat == 'downs' or stat == 'kills':
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present, 4)
			elif stat in config.buffs_stacking_duration:
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present, 4)
			elif stat in config.buffs_stacking_intensity:
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present, 4)
			else:
				player.average_stats[stat] = round(player.total_stats[stat]/player.duration_fights_present, 4)

				
	myprint(log, "\n")

	if anonymize:
		anonymize_players(players, account_index)
	
	return players, fights, found_healing, found_barrier, squad_comp, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats
			

# replace all acount names with "account <number>" and all player names with "anon <number>"
def anonymize_players(players, account_index):
	for account in account_index:
		for i in account_index[account]:
			players[i].account = "Account "+str(i)
	for i,player in enumerate(players):
		player.name = "Anon "+str(i)


def get_combat_start_from_player_json(initial_time, player_json):
	start_combat = -1
	# TODO check healthPercents exists
	last_health_percent = 100
	for change in player_json['healthPercents']:
		if change[0] < initial_time:
			last_health_percent = change[1]
			continue
		if change[1] - last_health_percent < 0:
			# got dmg
			start_combat = change[0]
			break
		last_health_percent = change[1]
	for i in range(math.ceil(initial_time/1000), len(player_json['damage1S'][0])):
		if i == 0:
			continue
		if player_json['powerDamage1S'][0][i] != player_json['powerDamage1S'][0][i-1]:
			if start_combat == -1:
				start_combat = i*1000
			else:
				start_combat = min(start_combat, i*1000)
			break
	return start_combat


def get_combat_time_breakpoints(player_json):
	start_combat = get_combat_start_from_player_json(0, player_json)
	if 'combatReplayData' not in player_json:
		print("WARNING: combatReplayData not in json, using activeTimes as time in combat")
		return [start_combat, get_stat_from_player_json(player_json, None, 'time_active', None) * 1000]
	replay = player_json['combatReplayData']
	if 'dead' not in replay:
		return [start_combat, get_stat_from_player_json(player_json, None, 'time_active', None) * 1000]

	breakpoints = []
	playerDeaths = dict(replay['dead'])
	playerDowns = dict(replay['down'])
	for deathKey, deathValue in playerDeaths.items():
		for downKey, downValue in playerDowns.items():
			if deathKey == downValue:
				if start_combat != -1:
					breakpoints.append([start_combat, deathKey])
				start_combat = get_combat_start_from_player_json(deathValue + 1000, player_json)
				break
	end_combat = (len(player_json['damage1S'][0]))*1000
	if start_combat != -1:
		breakpoints.append([start_combat, end_combat])

	return breakpoints

def sum_breakpoints(breakpoints):
	combat_time = 0
	for [start, end] in breakpoints:
		combat_time += end - start
	return combat_time

# get value of stat from player_json
def get_stat_from_player_json(player_json, players_running_healing_addon, stat, config, activeBuffs = False, buffGenType = BuffGenerationType.SQUAD):
	if stat == 'time_in_combat':
		return round(sum_breakpoints(get_combat_time_breakpoints(player_json)) / 1000)

	if stat == 'group':
		if 'group' not in player_json:
			return 0
		return int(player_json['group'])
	
	if stat == 'time_active':
		if 'activeTimes' not in player_json:
			return 0
		return round(int(player_json['activeTimes'][0])/1000)
	
	if stat == 'dmg_taken':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'damageTaken' not in player_json['defenses'][0] or 'damageBarrier' not in player_json['defenses'][0]:
			return 0
		return int(player_json['defenses'][0]['damageTaken'])

	if stat == 'barrierDamage':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'damageBarrier' not in player_json['defenses'][0]:
			return 0
		return int(player_json['defenses'][0]['damageBarrier'])
		
	if stat == 'deaths':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'deadCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['deadCount'])

	if stat == 'downed':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'downCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['downCount'])

	if stat == 'hitsMissed':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'missedCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['missedCount'])

	if stat == 'interupted':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'interruptedCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['interruptedCount'])

	if stat == 'dmg':
		if 'dpsTargets' not in player_json:
			return 0
		sumDamage = 0
		for target in player_json['dpsTargets']:
			sumDamage = sumDamage + int(target[0]['damage'])
		return int(sumDamage)
		#return int(player_json['dpsAll'][0]['damage'])            
	#Add Power and Condition Damage Tracking
	if stat == 'Cdmg':
		if 'dpsTargets' not in player_json:
			return 0
		sumDamage = 0
		for target in player_json['dpsTargets']:
			sumDamage = sumDamage + int(target[0]['condiDamage'])
		return int(sumDamage)
		#return int(player_json['dpsAll'][0]['condiDamage'])    
	
	if stat == 'Pdmg':
		if 'dpsTargets' not in player_json:
			return 0
		sumDamage = 0
		for target in player_json['dpsTargets']:
			sumDamage = sumDamage + int(target[0]['powerDamage'])
		return int(sumDamage)
		#return int(player_json['dpsAll'][0]['powerDamage'])  

	if stat == 'res':
		if 'support' not in player_json or len(player_json['support']) != 1 or 'resurrects' not in player_json['support'][0]:
			return 0
		return int(player_json['support'][0]['resurrects'])

	if stat == 'rips':
		if 'support' not in player_json or len(player_json['support']) != 1 or 'boonStrips' not in player_json['support'][0]:
			return 0
		return int(player_json['support'][0]['boonStrips'])
	
	if stat == 'cleanses':
		if 'support' not in player_json or len(player_json['support']) != 1 or 'condiCleanse' not in player_json['support'][0]:
			return 0
		return int(player_json['support'][0]['condiCleanse'])            
	#Prep work for new addition: incoming Boon Strips
	if stat == 'ripsIn':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'boonStrips' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['boonStrips'])

	if stat == 'ripsTime':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'boonStripsTime' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['boonStripsTime'])

	#Prep work for new addition: incoming Condition Clears		
	if stat == 'cleansesIn':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'conditionCleanses' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['conditionCleanses'])				

	if stat == 'cleansesTime':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'conditionCleansesTime' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['conditionCleansesTime'])		

	if stat == 'dodges':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'dodgeCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['dodgeCount'])	

	if stat == 'evades':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'evadedCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['evadedCount'])

	if stat == 'invulns':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'invulnedCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['invulnedCount'])

	if stat == 'blocks':
		if 'defenses' not in player_json or len(player_json['defenses']) != 1 or 'blockedCount' not in player_json['defenses'][0]:
			return 0        
		return int(player_json['defenses'][0]['blockedCount'])		

	if stat == 'dist':
		if 'statsAll' not in player_json or len(player_json['statsAll']) != 1 or 'distToCom' not in player_json['statsAll'][0]:
			return -1
		return float(player_json['statsAll'][0]['distToCom'])

	if stat == 'downContrib':
		if 'statsAll' not in player_json or len(player_json['statsAll']) != 1 or 'downContribution' not in player_json['statsAll'][0]:
			return -1
		return float(player_json['statsAll'][0]['downContribution'])

	if stat == 'swaps':
		if 'statsAll' not in player_json or len(player_json['statsAll']) != 1 or 'swapCount' not in player_json['statsAll'][0]:
			return -1
		return float(player_json['statsAll'][0]['swapCount'])

	if stat == 'kills':
		countKills = 0
		for target in player_json['statsTargets']:
			countKills = countKills + int(target[0]['killed'])
		return int(countKills)
		#if 'statsAll' not in player_json or len(player_json['statsAll']) != 1 or 'killed' not in player_json['statsAll'][0]:
		#    return -1
		#return float(player_json['statsAll'][0]['killed'])

	if stat == 'downs':
		countDowns = 0
		for target in player_json['statsTargets']:
			countDowns = countDowns + int(target[0]['downed'])
		return int(countDowns)
		#if 'statsAll' not in player_json or len(player_json['statsAll']) != 1 or 'downed' not in player_json['statsAll'][0]:
		#    return -1
		#return float(player_json['statsAll'][0]['downed'])

	### Buffs ###
	if stat in config.buff_ids:
		if buffGenType == BuffGenerationType.GROUP:
			buffArrayName = 'groupBuffs' if not activeBuffs else 'groupBuffsActive'
		elif buffGenType == BuffGenerationType.OFFGROUP:
			buffArrayName = 'offGroupBuffs' if not activeBuffs else 'offGroupBuffsActive'
		elif buffGenType == BuffGenerationType.SELF:
			buffArrayName = 'selfBuffs' if not activeBuffs else 'selfBuffsActive'
		else:
			buffArrayName = 'squadBuffs' if not activeBuffs else 'squadBuffsActive'

		if buffArrayName not in player_json:
			return 0
		# get buffs in squad generation -> need to loop over all buffs
		for buff in player_json[buffArrayName]:
			if 'id' not in buff:
				continue 
			# find right buff
			buffId = buff['id']
			if buffId == int(config.buff_ids[stat]):
				if 'generation' not in buff['buffData'][0]:
					return 0
				if stat == 'iol':
					return 1
				else:
					return float(buff['buffData'][0]['generation'])
		return 0

	if stat == 'heal':
		#if player_json['name'] in players_running_healing_addon and 'extHealingStats' in player_json and 'alliedHealing1S' in player_json['extHealingStats']:
		if player_json['name'] in players_running_healing_addon and 'extHealingStats' in player_json and 'outgoingHealingAllies' in player_json['extHealingStats']:
			playerHealing = 0
			for healingtarget in player_json['extHealingStats']['outgoingHealingAllies']:
				playerHealing += (healingtarget[0]['actorHealing']-healingtarget[0]['actorDownedHealing'])
			return playerHealing
			#return sum([healing[0][-1] for healing in player_json['extHealingStats']['alliedHealing1S']])
		return -1

	if stat == 'barrier':
		if player_json['name'] in players_running_healing_addon and 'extBarrierStats' in player_json and 'alliedBarrier1S' in player_json['extBarrierStats']:
			return sum([barrier[0][-1] for barrier in player_json['extBarrierStats']['alliedBarrier1S']])
		return -1

# DPS Stats block
arrow_cart_skill_ids = [18850, 18853, 18855, 18860, 18862, 18865, 18867, 18869, 18872]
trebuchet_skill_ids = [21037, 21038]
catapult_skill_ids = [20242, 20272]
cannon_skill_ids = [14626, 14658, 14659, 18535, 18531, 18543, 19626]
burning_oil_skill_ids = [18887]
dragon_banner_skill_ids = [32980, 31968, 33232]

siege_skill_ids = [
	*arrow_cart_skill_ids,
	*trebuchet_skill_ids,
	*catapult_skill_ids,
	*cannon_skill_ids,
	*burning_oil_skill_ids,
	*dragon_banner_skill_ids
]

def moving_average(data, window_size):
	num_elements = len(data)
	ma = []
	for n in range(num_elements):
		min_tick = max(n - window_size, 0)
		max_tick = min(n + window_size, num_elements - 1)
		sub_data = data[min_tick:max_tick + 1]
		ma.append(sum(sub_data) / len(sub_data))

	return ma

# States array is formatted: [start, stack_count]
# Reformat as: [start, end, stack_count]
def split_boon_states(states, duration):
	split_states = []
	num_states = len(states) - 1
	for index, [start, stacks] in enumerate(states):
		if index == num_states:
			if start < duration:
				split_states.append([start, duration, stacks])
		else:
			split_states.append([start, min(states[index + 1][0], duration), stacks])
	return split_states

# Take state array and combat breakpoints, filter down states to only include those when in combat
def split_boon_states_by_combat_breakpoints(states, breakpoints, duration):
	if not breakpoints:
		return []

	breakpoints_copy = breakpoints[:]
	split_states = split_boon_states(states, duration)
	new_states = []

	while(len(breakpoints_copy) > 0 and len(split_states) > 0):
		[combat_start, combat_end] = breakpoints_copy.pop(0)
		[start_state, end_state, stacks] = split_states.pop(0)

		while(end_state < combat_start):
			if len(split_states) == 0:
				break
			[start_state, end_state, stacks] = split_states.pop(0)

		if end_state < combat_start:
			break

		new_start = combat_start if combat_start > start_state else start_state
		new_end = combat_end if combat_end < end_state else end_state
		if new_end > new_start:
			new_states.append([new_start, new_end, stacks])

		while(len(split_states) > 0 and split_states[0][1] <= combat_end):
			[start_state, end_state, stacks] = split_states.pop(0)

			new_start = combat_start if combat_start > start_state else start_state
			new_end = combat_end if combat_end < end_state else end_state

			if new_end > new_start:
				new_states.append([
					combat_start if combat_start > start_state else start_state,
					combat_end if combat_end < end_state else end_state,
					stacks
				])

	return new_states

player_roles = {}
player_combat_time = {}
def calculate_dps_stats(fight_json, fight, players_running_healing_addon, config):
	if fight.skipped:
		return

	fight_ticks = len(fight_json['players'][0]["damage1S"][0])

	damagePS = {}
	for index, target in enumerate(fight_json['targets']):
		if 'enemyPlayer' in target and target['enemyPlayer'] == True:
			for player in fight_json['players']:
				player_prof_name  = "{{"+player['profession']+"}} "+player['name']		
				if player_prof_name  not in damagePS:
					damagePS[player_prof_name ] = [0] * fight_ticks

				damage_on_target = player["targetDamage1S"][index][0]
				for i in range(fight_ticks):
					damagePS[player_prof_name ][i] += damage_on_target[i]

	#player_roles = {}
	#player_combat_time = {}
	skip_fight = {}
	for player in fight_json['players']:
		player_prof_name = "{{"+player['profession']+"}} "+player['name']

		if player['notInSquad']:
			skip_fight[player_prof_name] = True
			continue

		time_in_combat = get_stat_from_player_json(player, players_running_healing_addon, 'time_in_combat', config)
		if time_in_combat == 0:
			skip_fight[player_prof_name] = True
			continue

		player_combat_time[player_prof_name] = time_in_combat
		player_roles[player_prof_name] = find_sub_type(player, time_in_combat)

		if 'dead' in player['combatReplayData'] and len(player['combatReplayData']['dead']) > 0 and (time_in_combat / fight.duration) < 0.4:
			skip_fight[player_prof_name] = True
		else:
			skip_fight[player_prof_name] = False

	squad_damage_per_tick = []
	for fight_tick in range(fight_ticks - 1):
		squad_damage_on_tick = 0
		for player in fight_json['players']:
			player_prof_name = "{{"+player['profession']+"}} "+player['name']
			if skip_fight[player_prof_name]:
				continue
	
			player_damage = damagePS[player_prof_name]
			squad_damage_on_tick += player_damage[fight_tick + 1] - player_damage[fight_tick]
		squad_damage_per_tick.append(squad_damage_on_tick)

	squad_damage_total = sum(squad_damage_per_tick)
	squad_damage_per_tick_ma = moving_average(squad_damage_per_tick, 1)
	squad_damage_ma_total = sum(squad_damage_per_tick_ma)

	CHUNK_DAMAGE_SECONDS = 21
	Ch5CaDamage1S = {}
	UsedOffensiveSiege = {}

	for player in fight_json['players']:
		player_prof_name = "{{"+player['profession']+"}} "+player['name']
		if skip_fight[player_prof_name]:
			continue

		player_role = player_roles[player_prof_name]
		DPSStats_prof_name = player_prof_name + " " + player_role	
		if DPSStats_prof_name not in DPSStats:
			DPSStats[DPSStats_prof_name] = {}
			DPSStats[DPSStats_prof_name]["account"] = player['account']
			DPSStats[DPSStats_prof_name]["name"] = player['name']
			DPSStats[DPSStats_prof_name]["profession"] = player['profession']
			DPSStats[DPSStats_prof_name]["role"] = player_role
			DPSStats[DPSStats_prof_name]["duration"] = 0
			DPSStats[DPSStats_prof_name]["combatTime"] = 0
			DPSStats[DPSStats_prof_name]["Coordination_Damage"] = 0
			DPSStats[DPSStats_prof_name]["Chunk_Damage"] = [0] * CHUNK_DAMAGE_SECONDS
			DPSStats[DPSStats_prof_name]["Chunk_Damage_Total"] = [0] * CHUNK_DAMAGE_SECONDS
			DPSStats[DPSStats_prof_name]["Carrion_Damage"] = 0
			DPSStats[DPSStats_prof_name]["Carrion_Damage_Total"] = 0
			DPSStats[DPSStats_prof_name]["Damage_Total"] = 0
			DPSStats[DPSStats_prof_name]["Squad_Damage_Total"] = 0
			DPSStats[DPSStats_prof_name]["Burst_Damage"] = [0] * CHUNK_DAMAGE_SECONDS
			DPSStats[DPSStats_prof_name]["Ch5Ca_Burst_Damage"] = [0] * CHUNK_DAMAGE_SECONDS
			DPSStats[DPSStats_prof_name]["Downs"] = 0
			DPSStats[DPSStats_prof_name]["Kills"] = 0
			
			
		Ch5CaDamage1S[player_prof_name] = [0] * fight_ticks
		UsedOffensiveSiege[player_prof_name] = False
			
		player_damage = damagePS[player_prof_name]
		
		DPSStats[DPSStats_prof_name]["duration"] += fight.duration
		DPSStats[DPSStats_prof_name]["combatTime"] += player_combat_time[player_prof_name]
		DPSStats[DPSStats_prof_name]["Damage_Total"] += player_damage[fight_ticks - 1]
		DPSStats[DPSStats_prof_name]["Squad_Damage_Total"] += squad_damage_total

		for statsTarget in player["statsTargets"]:
			DPSStats[DPSStats_prof_name]["Downs"] += statsTarget[0]['downed']
			DPSStats[DPSStats_prof_name]["Kills"] += statsTarget[0]['killed']

		for damage_dist in player['totalDamageDist'][0]:
			if damage_dist['id'] in siege_skill_ids:
				UsedOffensiveSiege[player_prof_name] = True

		if "minions" in player:	
			for minion in player["minions"]:
				for minion_damage_dist in minion["totalDamageDist"][0]:
					if minion_damage_dist['id'] in siege_skill_ids:
						UsedOffensiveSiege[player_prof_name] = True

		# Coordination_Damage: Damage weighted by coordination with squad
		player_damage_per_tick = [player_damage[0]]
		for fight_tick in range(fight_ticks - 1):
			player_damage_per_tick.append(player_damage[fight_tick + 1] - player_damage[fight_tick])

		player_damage_ma = moving_average(player_damage_per_tick, 1)

		for fight_tick in range(fight_ticks - 1):
			player_damage_on_tick = player_damage_ma[fight_tick]
			if player_damage_on_tick == 0:
				continue

			squad_damage_on_tick = squad_damage_per_tick_ma[fight_tick]
			if squad_damage_on_tick == 0:
				continue

			squad_damage_percent = squad_damage_on_tick / squad_damage_ma_total

			DPSStats[DPSStats_prof_name]["Coordination_Damage"] += player_damage_on_tick * squad_damage_percent * fight.duration

	# Chunk damage: Damage done within X seconds of target down
	for index, target in enumerate(fight_json['targets']):
		if 'enemyPlayer' in target and target['enemyPlayer'] == True and 'combatReplayData' in target and len(target['combatReplayData']['down']):
			for chunk_damage_seconds in range(1, CHUNK_DAMAGE_SECONDS):
				targetDowns = dict(target['combatReplayData']['down'])
				for targetDownsIndex, (downKey, downValue) in enumerate(targetDowns.items()):
					downIndex = math.ceil(downKey / 1000)
					startIndex = max(0, math.ceil(downKey / 1000) - chunk_damage_seconds)
					if targetDownsIndex > 0:
						lastDownKey, lastDownValue = list(targetDowns.items())[targetDownsIndex - 1]
						lastDownIndex = math.ceil(lastDownKey / 1000)
						if lastDownIndex == downIndex:
							# Probably an ele in mist form
							continue
						startIndex = max(startIndex, lastDownIndex)

					squad_damage_on_target = 0
					for player in fight_json['players']:
						player_prof_name = "{{"+player['profession']+"}} "+player['name']	
						if skip_fight[player_prof_name]:
							continue
						
						player_role = player_roles[player_prof_name]
						DPSStats_prof_name = player_prof_name + " " + player_role

						damage_on_target = player["targetDamage1S"][index][0]
						player_damage = damage_on_target[downIndex] - damage_on_target[startIndex]

						DPSStats[DPSStats_prof_name]["Chunk_Damage"][chunk_damage_seconds] += player_damage
						squad_damage_on_target += player_damage

						if chunk_damage_seconds == 5:
							for i in range(startIndex, downIndex):
								Ch5CaDamage1S[player_prof_name][i] += damage_on_target[i + 1] - damage_on_target[i]

					for player in fight_json['players']:
						player_prof_name = "{{"+player['profession']+"}} "+player['name']
						if skip_fight[player_prof_name]:
							continue

						player_role = player_roles[player_prof_name]
						DPSStats_prof_name = player_prof_name + " " + player_role

						DPSStats[DPSStats_prof_name]["Chunk_Damage_Total"][chunk_damage_seconds] += squad_damage_on_target

	# Carrion damage: damage to downs that die 
	for index, target in enumerate(fight_json['targets']):
		if 'enemyPlayer' in target and target['enemyPlayer'] == True and 'combatReplayData' in target and len(target['combatReplayData']['dead']):
			targetDeaths = dict(target['combatReplayData']['dead'])
			targetDowns = dict(target['combatReplayData']['down'])
			for deathKey, deathValue in targetDeaths.items():
				for downKey, downValue in targetDowns.items():
					if deathKey == downValue:
						dmgEnd = math.ceil(deathKey / 1000)
						dmgStart = math.ceil(downKey / 1000)

						total_carrion_damage = 0
						for player in fight_json['players']:
							player_prof_name = "{{"+player['profession']+"}} "+player['name']
							if skip_fight[player_prof_name]:
								continue
							
							player_role = player_roles[player_prof_name]
							DPSStats_prof_name = player_prof_name + " " + player_role
							damage_on_target = player["targetDamage1S"][index][0]
							carrion_damage = damage_on_target[dmgEnd] - damage_on_target[dmgStart]

							DPSStats[DPSStats_prof_name]["Carrion_Damage"] += carrion_damage
							total_carrion_damage += carrion_damage

							for i in range(dmgStart, dmgEnd):
								Ch5CaDamage1S[player_prof_name][i] += damage_on_target[i + 1] - damage_on_target[i]

						for player in fight_json['players']:
							player_prof_name = "{{"+player['profession']+"}} "+player['name']
							if skip_fight[player_prof_name]:
								continue
							
							player_role = player_roles[player_prof_name]
							DPSStats_prof_name = player_prof_name + " " + player_role
							DPSStats[DPSStats_prof_name]["Carrion_Damage_Total"] += total_carrion_damage

	# Burst damage: max damage done in n seconds
	for player in fight_json['players']:
		player_prof_name = "{{"+player['profession']+"}} "+player['name']
		if skip_fight[player_prof_name] or UsedOffensiveSiege[player_prof_name]:
			# Exclude Dragon Banner from Burst stats
			continue

		player_role = player_roles[player_prof_name]
		DPSStats_prof_name = player_prof_name + " " + player_role
		player_damage = damagePS[player_prof_name]
		for i in range(1, CHUNK_DAMAGE_SECONDS):
			for fight_tick in range(i, fight_ticks):
				dmg = player_damage[fight_tick] - player_damage[fight_tick - i]
				DPSStats[DPSStats_prof_name]["Burst_Damage"][i] = max(dmg, DPSStats[DPSStats_prof_name]["Burst_Damage"][i])

	# Ch5Ca Burst damage: max damage done in n seconds
	for player in fight_json['players']:
		player_prof_name = "{{"+player['profession']+"}} "+player['name']
		if skip_fight[player_prof_name] or UsedOffensiveSiege[player_prof_name]:
			# Exclude Dragon Banner from Burst stats
			continue

		player_role = player_roles[player_prof_name]
		DPSStats_prof_name = player_prof_name + " " + player_role
		player_damage_ps = Ch5CaDamage1S[player_prof_name]
		player_damage = [0] * len(player_damage_ps)
		player_damage[0] = player_damage_ps[0]
		for i in range(1, len(player_damage)):
			player_damage[i] = player_damage[i - 1] + player_damage_ps[i]
		for i in range(1, CHUNK_DAMAGE_SECONDS):
			for fight_tick in range(i, fight_ticks):
				dmg = player_damage[fight_tick] - player_damage[fight_tick - i]
				DPSStats[DPSStats_prof_name]["Ch5Ca_Burst_Damage"][i] = max(dmg, DPSStats[DPSStats_prof_name]["Ch5Ca_Burst_Damage"][i])
	
	# Track Stacking Buff Uptimes
	damage_with_buff_buffs = ['stability', 'protection', 'aegis', 'might', 'fury', 'resistance', 'resolution', 'quickness', 'swiftness', 'alacrity', 'vigor', 'regeneration']
	for player in fight_json['players']:
		player_prof_name = "{{"+player['profession']+"}} "+player['name']
		if skip_fight[player_prof_name]:
			continue

		player_role = player_roles[player_prof_name]
		DPSStats_prof_name = player_prof_name + " " + player_role
		if DPSStats_prof_name not in stacking_uptime_Table:
			stacking_uptime_Table[DPSStats_prof_name] = {}
			stacking_uptime_Table[DPSStats_prof_name]["account"] = player['account']
			stacking_uptime_Table[DPSStats_prof_name]["name"] = player['name']
			stacking_uptime_Table[DPSStats_prof_name]["profession"] = player['profession']
			stacking_uptime_Table[DPSStats_prof_name]["role"] = player_role
			stacking_uptime_Table[DPSStats_prof_name]["duration_might"] = 0
			stacking_uptime_Table[DPSStats_prof_name]["duration_stability"] = 0
			stacking_uptime_Table[DPSStats_prof_name]["might"] = [0] * 26
			stacking_uptime_Table[DPSStats_prof_name]["stability"] = [0] * 26
			for buff_name in damage_with_buff_buffs:
				stacking_uptime_Table[DPSStats_prof_name]["damage_with_"+buff_name] = [0] * 26 if buff_name == 'might' else [0] * 2

		player_damage = damagePS[player_prof_name]
		player_damage_per_tick = [player_damage[0]]
		for fight_tick in range(fight_ticks - 1):
			player_damage_per_tick.append(player_damage[fight_tick + 1] - player_damage[fight_tick])

		player_combat_breakpoints = get_combat_time_breakpoints(player)

		for item in player['buffUptimesActive']:
			buffId = int(item['id'])	
			if buffId not in uptime_Buff_Ids:
				continue

			buff_name = uptime_Buff_Ids[buffId]
			if buff_name in damage_with_buff_buffs:
				states = split_boon_states_by_combat_breakpoints(item['states'], player_combat_breakpoints, fight.duration*1000)

				total_time = 0
				for idx, [state_start, state_end, stacks] in enumerate(states):
					if buff_name in ['stability', 'might']:
						uptime = state_end - state_start
						total_time += uptime
						stacking_uptime_Table[DPSStats_prof_name][buff_name][min(stacks, 25)] += uptime

					if buff_name in damage_with_buff_buffs:
						start_sec = state_start / 1000
						end_sec = state_end / 1000

						start_sec_int = int(start_sec)
						start_sec_rem = start_sec - start_sec_int

						end_sec_int = int(end_sec)
						end_sec_rem = end_sec - end_sec_int

						damage_with_stacks = 0
						if start_sec_int == end_sec_int:
							damage_with_stacks = player_damage_per_tick[start_sec_int] * (end_sec - start_sec)
						else:
							damage_with_stacks = player_damage_per_tick[start_sec_int] * (1.0 - start_sec_rem)
							damage_with_stacks += sum(player_damage_per_tick[start_sec_int + s] for s in range(1, end_sec_int - start_sec_int))
							damage_with_stacks += player_damage_per_tick[end_sec_int] * end_sec_rem

						if idx == 0:
							# Get any damage before we have boon states
							damage_with_stacks += player_damage_per_tick[start_sec_int] * (start_sec_rem)
							damage_with_stacks += sum(player_damage_per_tick[s] for s in range(0, start_sec_int))
						if idx == len(states) - 1:
							# leave this as if, not elif, since we can have 1 state which is both the first and last
							# Get any damage after we have boon states
							damage_with_stacks += player_damage_per_tick[end_sec_int] * (1.0 - end_sec_rem)
							damage_with_stacks += sum(player_damage_per_tick[s] for s in range(end_sec_int + 1, len(player_damage_per_tick)))
						elif len(states) > 1 and state_end != states[idx + 1][0]:
							# Get any damage between deaths, this is usually a small amount of condis that are still ticking after death
							next_state_start = states[idx + 1][0]
							next_state_sec = next_state_start / 1000
							next_start_sec_int = int(next_state_sec)
							next_start_sec_rem = next_state_sec - next_start_sec_int

							damage_with_stacks += player_damage_per_tick[end_sec_int] * (1.0 - end_sec_rem)
							damage_with_stacks += sum(player_damage_per_tick[s] for s in range(end_sec_int + 1, next_start_sec_int))
							damage_with_stacks += player_damage_per_tick[next_start_sec_int] * (next_start_sec_rem)

						if buff_name == 'might':
							stacking_uptime_Table[DPSStats_prof_name]["damage_with_"+buff_name][min(stacks, 25)] += damage_with_stacks
						else:
							stacking_uptime_Table[DPSStats_prof_name]["damage_with_"+buff_name][min(stacks, 1)] += damage_with_stacks

				if buff_name in ['stability', 'might']:
					stacking_uptime_Table[DPSStats_prof_name]["duration_"+buff_name] += total_time

	return DPSStats

# get stats for this fight from fight_json
# Input:
# fight_json = json object including one fight
# config = the config to use
# log = log file to write to
def get_stats_from_fight_json(fight_json, config, log):
	# get fight duration
	fight_duration_json = fight_json['duration']
	#split_duration = fight_duration_json.split('m ', 1)
	#mins = int(split_duration[0])
	#split_duration = split_duration[1].split('s', 1)
	#secs = int(split_duration[0])
	#if debug:
	#	print("duration: ", mins, "m", secs, "s")
	#duration = mins*60 + secs
	duration = round(fight_json['durationMS']/1000)

	num_allies = len(fight_json['players'])
	num_enemies = 0
	enemy_name = ''
	enemy_squad = {}
	num_kills = 0
	num_downs = 0
	enemy_Dps = {}
	enemyDps_name = ''
	enemyDps_damage = 0
	enemy_skill_dmg = {}
	squad_skill_dmg = {}
	squad_Dps = {}
	squadDps_name = ''
	squadDps_profession = ''
	squadDps_damage = 0
	squad_spike_dmg = {}
	fight_name = fight_json['timeEnd'].split(" -",1)[0]
	squad_damage_output[fight_name] = {}

	#creat dictionary of skill_ids and skill_names
	skills = fight_json['skillMap']
	for skill_id, skill in skills.items():
		x_id=skill_id[1:]
		if x_id not in skill_Dict:
			skill_Dict[x_id] = {}
			skill_Dict[x_id]['name'] = skill['name']
			skill_Dict[x_id]['icon'] = skill['icon']
	skillBuffs = fight_json['buffMap']
	for skill_id, skill in skillBuffs.items():
		x_id=skill_id[1:]
		if x_id not in skill_Dict:
			skill_Dict[x_id] = {}
			skill_Dict[x_id]['name'] = skill['name']
			skill_Dict[x_id]['icon'] = skill['icon']

	SiegeSkills = {14627: "Punch", 14639: "Whirling Assualt", 14709: "Rocket Punch", 14710: "Whirling Inferno", 14708: "Rocket Salvo"}

	for enemy in fight_json['targets']:
		if 'enemyPlayer' in enemy and enemy['enemyPlayer'] == True:
			num_enemies += 1
			#append enemy['name'] to enemy_list
			enemy_name = enemy['name'].split(' pl')[0]
			enemyDps_name = "{{"+enemy_name+"}} "+enemy['name']
			enemyDps_damage = enemy['dpsAll'][0]['damage']
			enemy_Dps[enemyDps_name] = enemyDps_damage
			
			for skill_used in enemy['totalDamageDist'][0]:
				skill_id = skill_used['id']
				if str(skill_id) in SiegeSkills:
					continue
				if str(skill_id) in skill_Dict:
					#skill_name = skill_Dict[str(skill_id)]
					skill_name = skill_Dict[str(skill_id)]['name']
				else:
					skill_name = 'Skill-'+str(skill_id)
				#skill_name = skill_Dict[skill_id]
				skill_dmg = skill_used['totalDamage']
				if skill_name not in enemy_skill_dmg:
					enemy_skill_dmg[skill_name] = skill_dmg
				else:
					enemy_skill_dmg[skill_name] = enemy_skill_dmg[skill_name] +skill_dmg

			#Track MOA - Signet of Humility (Active)
			for skill in enemy['totalDamageTaken'][0]:
				if int(skill['id']) == 29519:
					if enemy_name not in MOA_Targets:
						MOA_Targets[enemy_name]={}
						MOA_Targets[enemy_name]['hits'] = skill['hits']
						MOA_Targets[enemy_name]['connectedHits'] = skill['connectedHits']
						MOA_Targets[enemy_name]['missed'] = skill['missed']
						MOA_Targets[enemy_name]['blocked'] = skill['blocked']
						MOA_Targets[enemy_name]['invulned'] = skill['invulned']
					else:
						MOA_Targets[enemy_name]['hits'] += skill['hits']
						MOA_Targets[enemy_name]['connectedHits'] += skill['connectedHits']
						MOA_Targets[enemy_name]['missed'] += skill['missed']
						MOA_Targets[enemy_name]['blocked'] += skill['blocked']
						MOA_Targets[enemy_name]['invulned'] += skill['invulned']
					

			#Tracking Outgoing Control Effects generated by the squad against enemy players
			Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
			#Control_Duration = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear',  26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
			#Control_Intensity = {833: 'Daze', 872: 'Stun'}
			for item in enemy['buffs']:
				conditionId = int(item['id'])
				if conditionId not in Control_Effects:
					continue
				skill_name = Control_Effects[conditionId]

				if skill_name not in squad_Control:
					squad_Control[skill_name] = {}
				for cc in item['buffData']:
					for key, value in cc['generated'].items():
						if key not in squad_Control[skill_name]:
							squad_Control[skill_name][key] = float((value/100)*duration)
						else:
							squad_Control[skill_name][key] = squad_Control[skill_name][key] + float((value/100)*duration)

			if enemy_name not in enemy_squad:
				enemy_squad[enemy_name] = 1
			else:
				enemy_squad[enemy_name] = enemy_squad[enemy_name] + 1
			
			if 'combatReplayData' in enemy:
				num_kills += len(enemy['combatReplayData']['dead'])
				num_downs += len(enemy['combatReplayData']['down'])

	for player in fight_json['players']:
		if player['notInSquad']:
			continue
		squadDps_name = player['name']
		squadDps_profession = player['profession']
		squadDps_prof_name = "{{"+squadDps_profession+"}} "+squadDps_name
		squadDps_damage = 0

		if squadDps_prof_name not in uptime_Table:
			uptime_Table[squadDps_prof_name]={}
			uptime_Table[squadDps_prof_name]['name']=squadDps_name
			uptime_Table[squadDps_prof_name]['prof']=squadDps_profession
			uptime_Table[squadDps_prof_name]['duration'] = 0
			print('Added player to uptime_Table: '+ squadDps_prof_name)		

		for target in player['dpsTargets']:
			squadDps_damage = squadDps_damage + int(target[0]['damage'])
		squad_Dps[squadDps_prof_name] = squadDps_damage

		for skill_used in player['totalDamageDist'][0]:
			skill_id = skill_used['id']
			if skill_id in SiegeSkills:
				continue            
			if str(skill_id) in skill_Dict:
				#skill_name = skill_Dict[str(skill_id)]
				skill_name = skill_Dict[str(skill_id)]['name']
			else:
				skill_name = 'Skill-'+str(skill_id)            
			skill_dmg = skill_used['totalDamage']
			if skill_name not in squad_skill_dmg:
				squad_skill_dmg[skill_name] = skill_dmg
			else:
				squad_skill_dmg[skill_name] = squad_skill_dmg[skill_name] +skill_dmg        
		#Collect Spike Damage for first 60 seconds of each fight
		sec_dmg = 0
		#fight_name = fight_json['timeEnd'].split(" -",1)[0]
		#squad_damage_output[fight_name] = {}
		for idx, damage in enumerate(player['damage1S'][0]):
			if damage > sec_dmg:
				if idx in squad_damage_output[fight_name]:
					squad_damage_output[fight_name][idx] += (damage-sec_dmg)
				else:
					squad_damage_output[fight_name][idx] = (damage-sec_dmg)
			sec_dmg = damage

		player_combat_time = sum_breakpoints(get_combat_time_breakpoints(player)) / 1000

		#Track MOA Activity - Casting of Signet of Humility
		if player['profession'] in ['Mesmer', 'Mirage', 'Chronomancer']:
			if 'rotation' in player:
				for item in player['rotation']:
					if int(item['id']) == 29519:
						if squadDps_name not in MOA_Casters:
							MOA_Casters[squadDps_name]={}
							MOA_Casters[squadDps_name]['attempts'] = len(item['skills'])
						else:
							MOA_Casters[squadDps_name]['attempts'] += len(item['skills'])

		#Track Incoming Control Effects generated by the enemy against Squad Members
		Control_Effects = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: "Hunter's Mark"}
		#config.condition_ids = {720: 'Blinded', 721: 'Crippled', 722: 'Chilled', 727: 'Immobile', 742: 'Weakness', 791: 'Fear', 833: 'Daze', 872: 'Stun', 26766: 'Slow', 27705: 'Taunt', 30778: 'Hunters Mark'}
		for item in player['buffUptimesActive']:
			conditionId = int(item['id'])
			if conditionId not in Control_Effects:
				continue
			skill_name = Control_Effects[conditionId]
			if skill_name not in enemy_Control:
				enemy_Control[skill_name] = {}
			if skill_name not in enemy_Control_Player:
				enemy_Control_Player[skill_name] = {}
			for cc in item['buffData']:
				for key, value in cc['generated'].items():
					if key not in enemy_Control_Player[skill_name]:
						enemy_Control_Player[skill_name][key] = float((value/100)*player_combat_time)
					else:
						enemy_Control_Player[skill_name][key] = enemy_Control_Player[skill_name][key] + float((value/100)*player_combat_time)
					if player['name'] not in enemy_Control[skill_name]:
						enemy_Control[skill_name][player['name']] = float((value/100)*player_combat_time)
					else:
						enemy_Control[skill_name][player['name']] = enemy_Control[skill_name][player['name']] + float((value/100)*player_combat_time)

		#Track Offensive stats from [statsTarets]
		statAll = ["totalDamageCount", "directDamageCount", "connectedDirectDamageCount", "connectedDamageCount", "critableDirectDamageCount", "criticalRate", "criticalDmg", "flankingRate", "againstMovingRate", "glanceRate", "missed", "evaded", "blocked", "interrupts", "invulned"]
		#squadDps_prof_name = player['name']
		#squadDps_profession = player['profession']
		#squadDps_prof_name = "{{"+squadDps_profession+"}} "+squadDps_name		

		if squadDps_prof_name not in squad_offensive:
			squad_offensive[squadDps_prof_name]={}
			squad_offensive[squadDps_prof_name]['name']= squadDps_name
			squad_offensive[squadDps_prof_name]['prof']= squadDps_profession
			squad_offensive[squadDps_prof_name]['stats']= {}
            
		for stat in statAll:
			if stat not in squad_offensive[squadDps_prof_name]['stats']:
				squad_offensive[squadDps_prof_name]['stats'][stat] = sum([stats[0][stat] for stats in player['statsTargets']])
			else:
				squad_offensive[squadDps_prof_name]['stats'][stat] += sum([stats[0][stat] for stats in player['statsTargets']])


		#Instant Revive tracking of downed healing
		instant_Revive = {14419: 'Battle Standard', 9163: 'Signet of Mercy', 5763: 'Renewal of Water', 5762: 'Renewal of Fire', 5760: 'Renewal of Air', 5761: 'Renewal of Earth', 10611: 'Signet of Undeath', 12596: "Nature's Renewal"}
		if 'extHealingStats' in player:
			for target in player['extHealingStats']['totalHealingDist'][0]:
				if 'totalDownedHealing' in target:
					if int(target['totalDownedHealing']) > 0:
						if target['id'] in instant_Revive:
							reviveSkill = instant_Revive[target['id']]

							if squadDps_prof_name not in downed_Healing:
								downed_Healing[squadDps_prof_name]={}
								downed_Healing[squadDps_prof_name]['name'] = squadDps_name
								downed_Healing[squadDps_prof_name]['prof'] = squadDps_profession
							if reviveSkill not in downed_Healing[squadDps_prof_name]:
								downed_Healing[squadDps_prof_name][reviveSkill] = {}
								downed_Healing[squadDps_prof_name][reviveSkill]['Heals'] = {}
								downed_Healing[squadDps_prof_name][reviveSkill]['Hits'] = {}
								downed_Healing[squadDps_prof_name][reviveSkill]['Heals'] = int(target['totalDownedHealing'])
								downed_Healing[squadDps_prof_name][reviveSkill]['Hits'] = int(target['hits'])
							else:
								downed_Healing[squadDps_prof_name][reviveSkill]['Heals'] = downed_Healing[squadDps_prof_name][reviveSkill]['Heals'] + int(target['totalDownedHealing'])
								downed_Healing[squadDps_prof_name][reviveSkill]['Hits'] = downed_Healing[squadDps_prof_name][reviveSkill]['Hits'] + int(target['hits'])	
		#End Instant Revive tracking
									
		#Track Aura Output		
		Auras_Id = {5677: 'Fire', 5577: 'Shocking', 5579: 'Frost', 5684: 'Magnetic', 25518: 'Light', 39978: 'Dark', 10332: 'Chaos'}
		for item in player['buffUptimesActive']:
			auraId = int(item['id'])
			if auraId not in Auras_Id:
				continue
			skill_name = Auras_Id[auraId]
			if skill_name not in auras_TableOut:
				auras_TableOut[skill_name] = {}				
			for cc in item['buffData']:
				for key, value in cc['generated'].items():
					if key not in auras_TableOut[skill_name]:
						auras_TableOut[skill_name][key] = float((value/100)*player_combat_time)
					else:
						auras_TableOut[skill_name][key] = auras_TableOut[skill_name][key] + float((value/100)*player_combat_time)

		#Track Total Buff Uptimes
		uptime_Buff_Ids = {1122: 'stability', 717: 'protection', 743: 'aegis', 740: 'might', 725: 'fury', 26980: 'resistance', 873: 'resolution', 1187: 'quickness', 719: 'swiftness', 30328: 'alacrity', 726: 'vigor', 718: 'regeneration'}
		#uptime_Buff_Names = { 'stability': 1122,  'protection': 717,  'aegis': 743,  'might': 740,  'fury': 725,  'resistance': 26980,  'resolution': 873,  'quickness': 1187,  'swiftness': 719,  'alacrity': 30328,  'vigor': 726,  'regeneration': 718}
		for item in player['buffUptimesActive']:
			buffId = int(item['id'])	
			if buffId not in uptime_Buff_Ids:
				continue
			buff_name = uptime_Buff_Ids[buffId]
			if buff_name == 'stability' or buff_name == 'might':
				uptime_value = float(item['buffData'][0]['presence'])
			else:
				uptime_value = float(item['buffData'][0]['uptime'])
			uptime_duration = float(player_combat_time * (uptime_value/100))
			if buff_name not in uptime_Table[squadDps_prof_name]:
				uptime_Table[squadDps_prof_name][buff_name] = uptime_duration
			else:
				uptime_Table[squadDps_prof_name][buff_name] = uptime_Table[squadDps_prof_name][buff_name] + uptime_duration
		uptime_Table[squadDps_prof_name]['duration'] = uptime_Table[squadDps_prof_name]['duration'] + player_combat_time

	#Attendance Tracking
	#duration in secs
	for player in fight_json['players']:
		player_account = player['account']
		player_name = player['name']
		player_prof = player['profession']

		if player['notInSquad']:
			continue
		
		if 'usedExtensions' not in fight_json:
			players_running_healing_addon = []
		else:
			extensions = fight_json['usedExtensions']
			for extension in extensions:
				if extension['name'] == "Healing Stats":
					players_running_healing_addon = extension['runningExtension']

		time_in_combat = get_stat_from_player_json(player, players_running_healing_addon, 'time_in_combat', config)

		if time_in_combat == 0:
			continue

		player_role = find_sub_type(player, time_in_combat)
		player_prof_role = player_prof+" "+player_role

		if Guild_Data:
			guildStatus = findMember(members, player_account)
		else:
			guildStatus = ""			

		if player_account not in Attendance:

			Attendance[player['account']]={}
			Attendance[player['account']]['fights'] = 1
			Attendance[player['account']]['duration'] = duration
			Attendance[player['account']]['guildStatus'] = guildStatus
			Attendance[player['account']]['names']={}
		
		elif player_account in Attendance:
			Attendance[player['account']]['fights'] += 1
			Attendance[player['account']]['duration'] += duration

		if player_name not in Attendance[player_account]['names']:
			Attendance[player_account]['names'][player_name]={}

			if player_prof_role not in Attendance[player_account]['names'][player_name]:
				Attendance[player_account]['names'][player_name]['professions']={}
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]={}
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['role'] = player_role
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['fights'] = 1
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['duration'] = duration
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['guildStatus'] = guildStatus
			else:
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['fights'] += 1
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['duration'] += duration
				Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['guildStatus'] = guildStatus


		#if player_name in Attendance[player_account]['names']:

		elif player_prof_role not in Attendance[player_account]['names'][player_name]['professions']:
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]={}
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['role'] = player_role
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['fights'] = 1
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['duration'] = duration
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['guildStatus'] = guildStatus
		else:
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['fights'] += 1
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['duration'] += duration
			Attendance[player_account]['names'][player_name]['professions'][player_prof_role]['guildStatus'] = guildStatus

	#Personal Buff Tracking
	personalBuffs = fight_json['personalBuffs']
	for prof in personalBuffs:
		if prof not in buffs_personal:
			buffs_personal[prof] = {}
			buffs_personal[prof]['buffList'] = []
			buffs_personal[prof]['player'] = {}

		for buff in personalBuffs[prof]:
			if buff not in buffs_personal[prof]['buffList']:
				buffs_personal[prof]['buffList'].append(buff)

	for player in fight_json['players']:
		player_prof = player['profession']
		player_name = player['name']
		player_activeTime = round(player['activeTimes'][0]/1000,2)
		if player_prof in buffs_personal:
			for buff in buffs_personal[player_prof]['buffList']:
				for activeBuff in player['buffUptimesActive']:
					if activeBuff['id'] == buff:
						buffUptime = activeBuff['buffData'][0]['uptime']
						uptimeSeconds =round(((buffUptime*player_activeTime)/100),2)
						if player_name not in buffs_personal[player_prof]['player']:
							buffs_personal[player_prof]['player'][player_name]={}
							buffs_personal[player_prof]['player'][player_name][buff]=0
							buffs_personal[player_prof]['player'][player_name][buff]=uptimeSeconds
						elif buff not in buffs_personal[player_prof]['player'][player_name]:
							buffs_personal[player_prof]['player'][player_name][buff]={}
							buffs_personal[player_prof]['player'][player_name][buff]=uptimeSeconds
						else:
							buffs_personal[player_prof]['player'][player_name][buff]+=uptimeSeconds



	#Death_OnTag Tracking
	tagPositions = {}
	dead_Tag = 0
	dead_Tag_Mark = 0
	commanderMissing = True
	commanderFound = False
	inchToPixel = fight_json['combatReplayMetaData']['inchToPixel']
	i=0
	for id in fight_json['players']:
		if id['hasCommanderTag']:
			commanderFound = True
			tagPositions = id['combatReplayData']['positions']
			if id['combatReplayData']['dead']:
				for death in id['combatReplayData']['dead']:
					dead_Tag_Mark = death[0]
					dead_Tag = 1
			else:
				dead_Tag_Mark = 999999999
				dead_Tag = 0

	if commanderFound:
			commanderMissing = False

	for id in fight_json['players']:
		playerDistances = []
		playerDistToTag = id['statsAll'][0]['distToCom']
		deathOnTag_name = id['name']
		deathOnTag_profession = id['profession']
		deathOnTag_prof_name = "{{"+deathOnTag_profession+"}} "+deathOnTag_name
		if deathOnTag_prof_name not in Death_OnTag:
			Death_OnTag[deathOnTag_prof_name] = {}
			Death_OnTag[deathOnTag_prof_name]["name"] = deathOnTag_name
			Death_OnTag[deathOnTag_prof_name]["profession"] = deathOnTag_profession
			Death_OnTag[deathOnTag_prof_name]["distToTag"] = []
			Death_OnTag[deathOnTag_prof_name]["On_Tag"] = 0
			Death_OnTag[deathOnTag_prof_name]["Off_Tag"] = 0
			Death_OnTag[deathOnTag_prof_name]["Run_Back"] = 0
			Death_OnTag[deathOnTag_prof_name]["After_Tag_Death"] = 0
			Death_OnTag[deathOnTag_prof_name]["Total"] = 0
			Death_OnTag[deathOnTag_prof_name]["Ranges"] = []
		if commanderMissing:
			continue
		if id['combatReplayData']['dead'] and id['combatReplayData']['down']:
			playerDeaths = dict(id['combatReplayData']['dead'])
			playerDowns = dict(id['combatReplayData']['down'])
			playerDistToTag = id['statsAll'][0]['distToCom']
			for deathKey, deathValue in playerDeaths.items():
				if deathKey < 0: #Handle death on the field before main squad combat log starts
					Death_OnTag[deathOnTag_prof_name]["Off_Tag"] = Death_OnTag[deathOnTag_prof_name]["Off_Tag"] + 1
				for downKey, downValue in playerDowns.items():
					if deathKey == downValue:
						#process data for downKey
						positionMark = int(downKey/150)
						positionDown = id['combatReplayData']['positions'][positionMark]
						x1 = positionDown[0]
						y1 = positionDown[1]
						x2 = tagPositions[positionMark][0]
						y2 = tagPositions[positionMark][1]
						deathDistance = math.sqrt((x1-x2)**2 + (y1-y2)**2)
						#deathRange = deathDistance/0.01
						deathRange = deathDistance/inchToPixel
						Death_OnTag[deathOnTag_prof_name]["Total"] = Death_OnTag[deathOnTag_prof_name]["Total"] + 1
						if int(downKey) > int(dead_Tag_Mark) and dead_Tag:
							Death_OnTag[deathOnTag_prof_name]["After_Tag_Death"] = Death_OnTag[deathOnTag_prof_name]["After_Tag_Death"] + 1
							continue
						if deathRange <= On_Tag:
							Death_OnTag[deathOnTag_prof_name]["On_Tag"] = Death_OnTag[deathOnTag_prof_name]["On_Tag"] + 1
						if deathRange > Run_Back:
							Death_OnTag[deathOnTag_prof_name]["Run_Back"] = Death_OnTag[deathOnTag_prof_name]["Run_Back"] + 1
						if deathRange > On_Tag and deathRange <= Run_Back:
							Death_OnTag[deathOnTag_prof_name]["Off_Tag"] = Death_OnTag[deathOnTag_prof_name]["Off_Tag"] + 1
							Death_OnTag[deathOnTag_prof_name]["Ranges"] += [deathRange]
				if deathValue:
					playerDeadPoll = int(deathValue/150)
					playerPositions = id['combatReplayData']['positions']
					for position,tagPosition in zip(playerPositions[:playerDeadPoll], tagPositions[:playerDeadPoll]):
						deltaX = position[0] - tagPosition[0]
						deltaY = position[1] - tagPosition[1]
						playerDistances.append(math.sqrt(deltaX * deltaX + deltaY * deltaY))
					playerDistToTag = (sum(playerDistances) / len(playerDistances))/inchToPixel
		Death_OnTag[deathOnTag_prof_name]["distToTag"].append(playerDistToTag)

	#Collect Box Plot DPS data by Profession, Prof_Name, Name, Acct
	durationMS = fight_json['durationMS']
	num_enemies = len(fight_json['targets'])
	num_allies = len(fight_json['players'])

	for player in fight_json['players']:
		if player['notInSquad']:
			continue		
		if durationMS < config.min_fight_duration or num_allies < config.min_allied_players or num_enemies < config.min_enemy_players:
			continue
		playerDPS = 0
		playerDamage = 0
		playerCPS = 0
		playerCleanses = 0
		playerSPS = 0
		playerStrips = 0
		playerHPS = 0
		playerHeals = 0						
		name = player['name']
		acct = player['account']
		sub_type = player['profession'] + "_" + find_sub_type(player, durationMS / 1000)
		prof = sub_type
		prof_name = sub_type+"\n"+name
		for target in player['dpsTargets']:
			playerDamage += target[0]['damage']
			
		playerDPS = round(playerDamage/(durationMS/1000), 4)

		if playerDPS > 0:
			if prof_name not in DPS_List['prof_name']:
				DPS_List['prof_name'][prof_name] = []
			if prof not in DPS_List['prof']:
				DPS_List['prof'][prof] = []            
			if name not in DPS_List['name']:
				DPS_List['name'][name] = []
			if acct not in DPS_List['acct']:
				DPS_List['acct'][acct] = []
		if playerDPS > 0:
			DPS_List['acct'][acct].append(playerDPS)
			DPS_List['name'][name].append(playerDPS)
			DPS_List['prof_name'][prof_name].append(playerDPS)
			DPS_List['prof'][prof].append(playerDPS)
#End DPS Box Plot Data Collection

		#Collect Box Plot CPS data by Profession, Prof_Name, Name, Acct
		if 'support' not in player or len(player['support']) != 1 or 'condiCleanse' not in player['support'][0]:
			playerCleanses = 0
		else:
			playerCleanses += int(player['support'][0]['condiCleanse'])   
		
		playerCPS = round(playerCleanses/(durationMS/1000), 4)

		if playerCPS > 0:
			if prof_name not in CPS_List['prof_name']:
				CPS_List['prof_name'][prof_name] = []
			if prof not in CPS_List['prof']:
				CPS_List['prof'][prof] = []            
			if name not in CPS_List['name']:
				CPS_List['name'][name] = []
			if acct not in CPS_List['acct']:
				CPS_List['acct'][acct] = []
		if playerCPS > 0:
			CPS_List['acct'][acct].append(playerCPS)
			CPS_List['name'][name].append(playerCPS)
			CPS_List['prof_name'][prof_name].append(playerCPS)
			CPS_List['prof'][prof].append(playerCPS)
		#End CPS Box Plot Data Collection

		#Collect Box Plot SPS data by Profession, Prof_Name, Name, Acct
		if 'support' not in player or len(player['support']) != 1 or 'boonStrips' not in player['support'][0]:
			playerStrips = 0
		else:
			playerStrips += int(player['support'][0]['boonStrips'])   
		
		playerSPS = round(playerStrips/(durationMS/1000), 4)

		if playerSPS > 0:
			if prof_name not in SPS_List['prof_name']:
				SPS_List['prof_name'][prof_name] = []
			if prof not in SPS_List['prof']:
				SPS_List['prof'][prof] = []            
			if name not in SPS_List['name']:
				SPS_List['name'][name] = []
			if acct not in SPS_List['acct']:
				SPS_List['acct'][acct] = []
		if playerSPS > 0:
			SPS_List['acct'][acct].append(playerSPS)
			SPS_List['name'][name].append(playerSPS)
			SPS_List['prof_name'][prof_name].append(playerSPS)
			SPS_List['prof'][prof].append(playerSPS)
		#End SPS Box Plot Data Collection

		#Collect Box Plot HPS data by Profession, Prof_Name, Name, Acct
		if 'extHealingStats' in player:
			if 'outgoingHealingAllies' not in player['extHealingStats']:
				playerHeals = 0
			for outgoing_healing_json in player['extHealingStats']['outgoingHealingAllies']:
				for outgoing_healing_json2 in outgoing_healing_json:
					if 'healing' in outgoing_healing_json2:
						playerHeals += int(outgoing_healing_json2['healing'])
		
		playerHPS = round(playerHeals/(durationMS/1000), 4)

		if playerHPS > 0:
			if prof_name not in HPS_List['prof_name']:
				HPS_List['prof_name'][prof_name] = []
			if prof not in HPS_List['prof']:
				HPS_List['prof'][prof] = []            
			if name not in HPS_List['name']:
				HPS_List['name'][name] = []
			if acct not in HPS_List['acct']:
				HPS_List['acct'][acct] = []
		if playerHPS > 0:
			HPS_List['acct'][acct].append(playerHPS)
			HPS_List['name'][name].append(playerHPS)
			HPS_List['prof_name'][prof_name].append(playerHPS)
			HPS_List['prof'][prof].append(playerHPS)
		#End HPS Box Plot Data Collection

	# initialize fight         
	fight = Fight()
	fight.duration = duration
	fight.enemies = num_enemies
	fight.enemy_squad = enemy_squad
	fight.enemy_Dps = enemy_Dps
	fight.squad_Dps = squad_Dps
	fight.squad_spike_dmg = squad_spike_dmg
	fight.enemy_skill_dmg = enemy_skill_dmg
	fight.squad_skill_dmg = squad_skill_dmg
	fight.allies = num_allies
	fight.kills = num_kills
	fight.downs = num_downs
	fight.start_time = fight_json['timeStartStd']
	fight.end_time = fight_json['timeEndStd']        
	fight.total_stats = {key: 0 for key in config.stats_to_compute}
			
	# skip fights that last less than min_fight_duration seconds
	if(duration < config.min_fight_duration):
		fight.skipped = True
		print_string = "\nFight only took "+str(duration)+"s. Skipping fight."
		myprint(log, print_string)
		
	# skip fights with less than min_allied_players allies
	if num_allies < config.min_allied_players:
		fight.skipped = True
		print_string = "\nOnly "+str(num_allies)+" allied players involved. Skipping fight."
		myprint(log, print_string)

	# skip fights with less than min_enemy_players enemies
	if num_enemies < config.min_enemy_players:
		fight.skipped = True
		print_string = "\nOnly "+str(num_enemies)+" enemies involved. Skipping fight."
		myprint(log, print_string)

	if 'usedExtensions' not in fight_json:
		players_running_healing_addon = []
	else:
		extensions = fight_json['usedExtensions']
		for extension in extensions:
			if extension['name'] == "Healing Stats":
				players_running_healing_addon = extension['runningExtension']

	calculate_dps_stats(fight_json, fight, players_running_healing_addon, config)
		
	return fight, players_running_healing_addon, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats



# add up total stats over all fights
def get_overall_squad_stats(fights, config):
	# overall stats over whole squad
	overall_squad_stats = {key: 0 for key in config.stats_to_compute}
	for fight in fights:
		if not fight.skipped:
			for stat in config.stats_to_compute:
				overall_squad_stats[stat] += fight.total_stats[stat]
	return overall_squad_stats

def get_overall_raid_stats(fights):
	overall_raid_stats = {}
	used_fights = [f for f in fights if not f.skipped]

	overall_raid_stats['num_used_fights'] = len([f for f in fights if not f.skipped])
	overall_raid_stats['used_fights_duration'] = sum([f.duration for f in used_fights])
	overall_raid_stats['date'] = min([f.start_time.split()[0] for f in used_fights])
	overall_raid_stats['start_time'] = min([f.start_time.split()[1] for f in used_fights]) +"<br>,,UTC "+ used_fights[0].start_time.split()[2]+",,"
	overall_raid_stats['end_time'] = max([f.end_time.split()[1] for f in used_fights]) +"<br>,,UTC "+ used_fights[0].end_time.split()[2]+",,"
	overall_raid_stats['num_skipped_fights'] = len([f for f in fights if f.skipped])
	overall_raid_stats['min_allies'] = min([f.allies for f in used_fights])
	overall_raid_stats['max_allies'] = max([f.allies for f in used_fights])    
	overall_raid_stats['mean_allies'] = round(sum([f.allies for f in used_fights])/len(used_fights), 1)
	overall_raid_stats['min_enemies'] = min([f.enemies for f in used_fights])
	overall_raid_stats['max_enemies'] = max([f.enemies for f in used_fights])        
	overall_raid_stats['mean_enemies'] = round(sum([f.enemies for f in used_fights])/len(used_fights), 1)
	overall_raid_stats['total_downs'] = sum([f.downs for f in used_fights])
	overall_raid_stats['total_kills'] = sum([f.kills for f in used_fights])    
	return overall_raid_stats


# print the overall squad stats
def print_total_squad_stats(fights, overall_squad_stats, overall_raid_stats, found_healing, found_barrier, config, output):
	#get total duration in h, m, s
	total_fight_duration = {}
	total_fight_duration["h"] = int(overall_raid_stats['used_fights_duration']/3600)
	total_fight_duration["m"] = int((overall_raid_stats['used_fights_duration'] - total_fight_duration["h"]*3600) / 60)
	total_fight_duration["s"] = int(overall_raid_stats['used_fights_duration'] - total_fight_duration["h"]*3600 -  total_fight_duration["m"]*60)
	
	print_string = "The following stats are computed over "+str(overall_raid_stats['num_used_fights'])+" out of "+str(len(fights))+" fights.\n"
	myprint(output, print_string)
	
	# print total squad stats
	print_string = "Squad overall"
	i = 0
	printed_kills = False
	for stat in config.stats_to_compute:
		if stat in exclude_Stat:
			continue
		
		if i == 0:
			print_string += " "
		elif i == len(config.stats_to_compute)-1 and printed_kills:
			print_string += ", and "
		else:
			print_string += ", "
		i += 1
		#JEL - modified select outputs to utilize my_value() formatting     
		if stat == 'dmg':
			print_string += "did "+my_value(round(overall_squad_stats['dmg']))+" damage"
		elif stat == 'rips':
			print_string += "ripped "+my_value(round(overall_squad_stats['rips']))+" boons"
		elif stat == 'cleanses':
			print_string += "cleansed "+my_value(round(overall_squad_stats['cleanses']))+" conditions"
		elif stat == 'iol':
			print_string += "Illusioned "+my_value(round(overall_squad_stats['iol']))+" downs"
		elif stat in config.buff_ids and stat != 'iol':
			total_buff_duration = {}
			total_buff_duration["h"] = int(overall_squad_stats[stat]/3600)
			total_buff_duration["m"] = int((overall_squad_stats[stat] - total_buff_duration["h"]*3600)/60)
			total_buff_duration["s"] = int(overall_squad_stats[stat] - total_buff_duration["h"]*3600 - total_buff_duration["m"]*60)    
			
			print_string += "generated "
			if total_buff_duration["h"] > 0:
				print_string += str(total_buff_duration["h"])+"h "
			print_string += str(total_buff_duration["m"])+"m "+str(total_buff_duration["s"])+"s of "+stat
		elif stat == 'heal' and found_healing:
			print_string += "healed for "+str(my_value(round(overall_squad_stats['heal'])))
		elif stat == 'barrier' and found_barrier:
			print_string += "generated "+str(my_value(round(overall_squad_stats['barrier'])))+" barrier"
		elif stat == 'dmg_taken':
			print_string += "took "+my_value(round(overall_squad_stats['dmg_taken']))+" damage"
		elif stat == 'deaths':
			print_string += "killed "+str(overall_raid_stats['total_kills'])+" enemies and had "+str(round(overall_squad_stats['deaths']))+" deaths"
			printed_kills = True

	if not printed_kills:
		print_string += ", and killed "+str(overall_raid_stats['total_kills'])+" enemies"
	print_string += " over a total time of "
	if total_fight_duration["h"] > 0:
		print_string += str(total_fight_duration["h"])+"h "
	print_string += str(total_fight_duration["m"])+"m "+str(total_fight_duration["s"])+"s in "+str(overall_raid_stats['num_used_fights'])+" fights.\n"
	#JEL - Added Kill Death Ratio
	try:
		Raid_KDR = (round((overall_raid_stats['total_kills']/(round(overall_squad_stats['deaths']))),2))
	except:
		Raid_KDR = overall_raid_stats['total_kills']

	print_string += "\nKill Death Ratio for the session was ''"+str(Raid_KDR)+"''.\n"
	#JEL - Added beginning newline for TW5 spacing
	print_string += "\nThere were between "+str(overall_raid_stats['min_allies'])+" and "+str(overall_raid_stats['max_allies'])+" allied players involved (average "+str(round(overall_raid_stats['mean_allies'], 1))+" players).\n"
	print_string += "\nThe squad faced between "+str(overall_raid_stats['min_enemies'])+" and "+str(overall_raid_stats['max_enemies'])+" enemy players (average "+str(round(overall_raid_stats['mean_enemies'], 1))+" players).\n"    
		
	myprint(output, print_string)
	return total_fight_duration


# Write xls fight overview
# Input:
# fights = list of Fights
# overall_squad_stats = overall stats of the whole squad
# xls_output_filename = where to write to
def write_fights_overview_xls(fights, overall_squad_stats, overall_raid_stats, config, xls_output_filename):
	book = xlrd.open_workbook(xls_output_filename)
	wb = copy(book)
	if len(book.sheet_names()) == 0 or book.sheet_names()[0] != 'fights overview':
		print("Sheet 'fights overview' is not the first sheet in"+xls_output_filename+". Skippping fights overview.")
		return
	sheet1 = wb.get_sheet(0)

	sheet1.write(0, 1, "#")
	sheet1.write(0, 2, "Date")
	sheet1.write(0, 3, "Start Time")
	sheet1.write(0, 4, "End Time")
	sheet1.write(0, 5, "Duration in s")
	sheet1.write(0, 6, "Skipped")
	sheet1.write(0, 7, "Num. Allies")
	sheet1.write(0, 8, "Num. Enemies")
	sheet1.write(0, 9, "Kills")
	
	for i,stat in enumerate(config.stats_to_compute):
		sheet1.write(0, 10+i, config.stat_names[stat])

	for i,fight in enumerate(fights):
		skipped_str = "yes" if fight.skipped else "no"
		sheet1.write(i+1, 1, i+1)
		sheet1.write(i+1, 2, fight.start_time.split()[0])
		sheet1.write(i+1, 3, fight.start_time.split()[1])
		sheet1.write(i+1, 4, fight.end_time.split()[1])
		sheet1.write(i+1, 5, fight.duration)
		sheet1.write(i+1, 6, skipped_str)
		sheet1.write(i+1, 7, fight.allies)
		sheet1.write(i+1, 8, fight.enemies)
		sheet1.write(i+1, 9, fight.kills)
		for j,stat in enumerate(config.stats_to_compute):
			sheet1.write(i+1, 10+j, fight.total_stats[stat])

	#used_fights = [f for f in fights if not f.skipped]
	#used_fights_duration = sum([f.duration for f in used_fights])
	#num_used_fights = len(used_fights)
	#date = min([f.start_time.split()[0] for f in used_fights])
	#start_time = min([f.start_time.split()[1] for f in used_fights])
	#end_time = max([f.end_time.split()[1] for f in used_fights])
	#skipped_fights = len(fights) - num_used_fights
	#mean_allies = round(sum([f.allies for f in used_fights])/num_used_fights, 1)
	#mean_enemies = round(sum([f.enemies for f in used_fights])/num_used_fights, 1)
	#total_kills = sum([f.kills for f in used_fights])

	sheet1.write(len(fights)+1, 0, "Sum/Avg. in used fights")
	sheet1.write(len(fights)+1, 1, overall_raid_stats['num_used_fights'])
	sheet1.write(len(fights)+1, 2, overall_raid_stats['date'])
	sheet1.write(len(fights)+1, 3, overall_raid_stats['start_time'])
	sheet1.write(len(fights)+1, 4, overall_raid_stats['end_time'])    
	sheet1.write(len(fights)+1, 5, overall_raid_stats['used_fights_duration'])
	sheet1.write(len(fights)+1, 6, overall_raid_stats['num_skipped_fights'])
	sheet1.write(len(fights)+1, 7, overall_raid_stats['mean_allies'])    
	sheet1.write(len(fights)+1, 8, overall_raid_stats['mean_enemies'])
	sheet1.write(len(fights)+1, 9, overall_raid_stats['total_kills'])
	for i,stat in enumerate(config.stats_to_compute):
		sheet1.write(len(fights)+1, 10+i, overall_squad_stats[stat])

	wb.save(xls_output_filename)

#JEL - TW5 tweaks for markdown table output
def print_fights_overview(fights, overall_squad_stats, overall_raid_stats, config, output):
	stat_len = {}
	print_string = '<div style="overflow-x:auto;">\n'
	myprint(output, print_string)
	print_string = "|thead-dark table-hover w-auto scrollable|k"
	myprint(output, print_string)
	print_string = "| Fight # | Date | Ending | Secs | Skip | Allies | Enemies | Downs | Kills |"
	for stat in overall_squad_stats:
		if stat not in exclude_Stat:
			stat_len[stat] = max(len(config.stat_names[stat]), len(str(overall_squad_stats[stat])))
			print_string += " {{"+config.stat_names[stat]+"}} |"
	print_string += "h"
	myprint(output, print_string)
	#Only write overall raid stats summary for monthly output
	if config.include_comp_and_review:
		for i in range(len(fights)):
			fight = fights[i]
			skipped_str = "yes" if fight.skipped else "no"
			date = fight.start_time.split()[0]
			end_time = fight.end_time.split()[1]        
			print_string = "| "+str((i+1))+" | "+str(date)+" | "+str(end_time)+" | "+str(fight.duration)+" | "+skipped_str+" | "+str(fight.allies)+" | "+str(fight.enemies)+" | "+str(fight.downs)+" | "+str(fight.kills)+" |"
			for stat in overall_squad_stats:
				if stat not in exclude_Stat:
					print_string += " "+my_value(round(fight.total_stats[stat]))+"|"
			myprint(output, print_string)

	#print_string = f"| {overall_raid_stats['num_used_fights']:>3}"+" | "+f"{overall_raid_stats['date']:>7}"+" | "+f"{overall_raid_stats['start_time']:>10}"+" | "+f"{overall_raid_stats['end_time']:>8}"+" | "+f"{overall_raid_stats['used_fights_duration']:>13}"+" | "+f"{overall_raid_stats['num_skipped_fights']:>7}" +" | "+f"{round(overall_raid_stats['mean_allies']):>11}"+" | "+f"{round(overall_raid_stats['mean_enemies']):>12}"+" | "+f"{round(overall_raid_stats['total_downs']):>5}"+" | "+f"{overall_raid_stats['total_kills']:>5} |"
	print_string = f"| {overall_raid_stats['num_used_fights']:>3}"+" | "+f"{overall_raid_stats['date']:>7}"+" | "+f"{overall_raid_stats['end_time']:>8}"+" | "+f"{overall_raid_stats['used_fights_duration']:>13}"+" | "+f"{overall_raid_stats['num_skipped_fights']:>7}" +" | "+f"{round(overall_raid_stats['mean_allies']):>11}"+" | "+f"{round(overall_raid_stats['mean_enemies']):>12}"+" | "+f"{round(overall_raid_stats['total_downs']):>5}"+" | "+f"{overall_raid_stats['total_kills']:>5} |"
	for stat in overall_squad_stats:
		if stat not in exclude_Stat:
			print_string += " "+my_value(round(overall_squad_stats[stat]))+"|"
	print_string += "f\n\n"
	myprint(output, print_string)
	print_string = '\n\n</div>'
	myprint(output, print_string)

#JEL - write TW5 stat Chart tids
def write_stats_chart(players, top_players, stat, myDate, input_directory, config):
	#args.input_directory+"/
	stat_Name = config.stat_names[stat]
	fileDate = myDate
	fileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_"+stat+"_TW5_Chart.tid"
	chart_Output = open(fileTid, "w",encoding="utf-8")
	minStatSec= 1000
	maxStatSec = 0
	
	print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
	print_string +="\ncreator: Drevarr\n"
	print_string +="tags: ChartData\n"
	print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData\n'
	print_string +="type: application/javascript\n\n\n"

	print_string += "option = {\n\tlegend: {},\n\tgrid: {left: '20%'},\n\ttooltip: {},\n\tdataset: [\n\t\t{\n\t\tsource: [\n\t\t\t["
	
	if stat == 'deaths' or stat == 'kills' or stat == 'downs':
		print_string += "'"+stat_Name+"/Min',"
	else:
		print_string += "'"+stat_Name+"/Sec',"

	print_string += "'Total "+stat_Name+"', 'Name', 'Profession', 'Fights', 'Duration' ],\n"
	
	for i in reversed(range(len(top_players))):
		player = players[top_players[i]]
		if stat == 'kills' or stat == 'downs':
			print_string += "\t\t\t["+str(round(player.average_stats[stat]*60, 4))+", "+str(round(player.total_stats[stat]))+", '"+player.name+"', '{{"+player.profession+"}}', '"+str(player.num_fights_present)+"', '"+str(player.duration_fights_present)+"'"
		else: 
			print_string += "\t\t\t["+str(player.average_stats[stat])+", "+str(round(player.total_stats[stat]))+", '"+player.name+"', '{{"+player.profession+"}}', '"+str(player.num_fights_present)+"', '"+str(player.duration_fights_present)+"'"
		if i >= len(top_players):
			print_string +="]\n"
		else:
			print_string +="],\n"
		#Set minStatSec and maxStatSec
		if stat == 'kills' or stat == 'downs':
			if (player.average_stats[stat]*60) > maxStatSec:
				maxStatSec = round((player.average_stats[stat]*60), 4)
			if (player.average_stats[stat]*60) < minStatSec:
				minStatSec = round((player.average_stats[stat]*60), 4)
		else:
			if player.average_stats[stat] < minStatSec:
				minStatSec = player.average_stats[stat]
			if player.average_stats[stat] > maxStatSec:
				maxStatSec = player.average_stats[stat]


	print_string += "\t\t\t]\n\t\t},\n\t],\n"
	print_string += "    xAxis: [\n"
	print_string += "    {},\n"
	print_string += "  ],\n"
	print_string += "  yAxis: { \n"
	print_string += "    type: 'category'\n"
	print_string += "  },\n"
	print_string += "  visualMap: {\n"
	print_string += "    orient: 'horizontal',\n"
	print_string += "    left: 'center',\n"
	print_string += "    min: "+str(minStatSec)+",\n"
	print_string += "    max: "+str(maxStatSec)+",\n"
	print_string += "    precision: 2,\n"
	print_string += "    text: ['High "+stat_Name+" /Sec', 'Low "+stat_Name+"/Sec'],\n"
	print_string += "    // Map the score column to color\n"
	print_string += "    dimension: 0,\n"
	print_string += "    inRange: {\n"
	print_string += "      color: ['#FD665F', '#FFCE34', '#65B581']\n"
	print_string += "    }\n"
	print_string += "  },\n"
	print_string += "  series: [\n"
	print_string += "    {\n"
	print_string += "      name: 'Total "+stat_Name+"',\n"
	print_string += "      type: 'bar',\n"
	print_string += "      encode: {\n"
	print_string += "        // Map 'Total Stat' column to x-axis.\n"
	print_string += "        x: 'Total "+stat_Name+"',\n"
	print_string += "        // Map 'Name' row to y-axis.\n"
	print_string += "        y: 'Name',\n"
	print_string += "        tooltip: [3, 4, 5, 1, 0]\n"
	print_string += "      }\n"
	print_string += "    },\n"
	print_string += "    {\n"
	print_string += "      name: '"+stat_Name+"/Sec',\n"
	print_string += "      type: 'bar',\n"
	print_string += "      encode: {\n"
	print_string += "        // Map 'Total Stat' column to x-axis.\n"
	print_string += "        x: 'Total "+stat_Name+"/Sec',\n"
	print_string += "        // Map 'Name' row to y-axis.\n"
	print_string += "        y: 'Name',\n"
	print_string += "        tooltip: [3, 4, 5, 1, 0]\n"
	print_string += "      }\n"
	print_string += "    }\n"
	print_string += "  ]\n"
	print_string += "};\n"

	myprint(chart_Output, print_string)

	chart_Output.close()
# 	end write TW5 Chart tids

def write_stats_box_plots(players, top_players, stat, ProfessionColor, myDate, input_directory, config):
	#args.input_directory+"/
	stat_Name = config.stat_names[stat]
	fileDate = myDate
	fileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_"+stat+"_TW5_Chart.tid"
	chart_Output = open(fileTid, "w",encoding="utf-8")
	statBoxPlot_names = []
	statBoxPlot_profs = []
	statBoxPlot_data = []	
	chart_per_fight = ['res', 'kills', 'downs', 'swaps', 'dist', 'downContrib', 'hitsMissed', 'interupted', 'invulns', 'evades', 'blocks', 'dodges', 'downed', 'deaths', 'dmg_taken', 'barrierDamage', 'superspeed']
	chart_per_second = ['dmg', 'Pdmg', 'Cdmg', 'rips', 'cleanses', 'heal', 'barrier', 'ripsIn', 'cleansesIn'] 
	for i in reversed(range(len(top_players))):
		player= players[top_players[i]]
		statBoxPlot_names.append(player.name)
		statBoxPlot_profs.append(player.profession)
		statPerFight = []
		#/100.*fight.duration*(fight.allies - 1)
		for fight in player.stats_per_fight:
			if fight[stat] != -1:
				duration = fight['fight_duration']
				fightAllies = fight['allies']
				if stat in chart_per_fight:
					statPerFight.append(round(fight[stat], 4))
				elif stat in chart_per_second:
					statPerFight.append(round(fight[stat]/duration, 4))
				else:
					statPerFight.append(round(fight[stat]/100*duration*(fightAllies-1),4))
        
		statBoxPlot_data.append(statPerFight)


	print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
	print_string +="\ncreator: Drevarr\n"
	print_string +="tags: ChartData\n"
	print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_'+stat+'_ChartData\n'
	print_string +="type: application/javascript\n\n\n"
	#output Box Plot Names
	jsonStr = json.dumps(statBoxPlot_names)
	print_string +='const names = '+jsonStr+';\n'
	#output Box Plot Professions
	jsonStr = json.dumps(statBoxPlot_profs)
	print_string +='const professions = '+jsonStr+';\n'
	#output Box Plot Professions
	jsonStr = json.dumps(ProfessionColor)
	print_string +='const ProfessionColor = '+jsonStr+';\n'
	
	print_string += "option = {\n"
	print_string += "  title: [\n"
	if stat in chart_per_fight:
			print_string += "    {text: '"+stat_Name+" per Fight for all Fights Present', left: 'center'},\n"
	elif stat in chart_per_second:
			print_string += "    {text: '"+stat_Name+" per Second for all Fights Present', left: 'center'},\n"
	else:
			print_string += "    {text: '"+stat_Name+" per Fight for all Fights Present', left: 'center'},\n"
	print_string += "    {text: 'Output in seconds across all fights \\nupper: Q3 + 1.5 * IQR \\nlower: Q1 - 1.5 * IQR', borderColor: '#999', borderWidth: 1, textStyle: {fontSize: 10}, left: '1%', top: '90%'}\n"
	print_string += "  ],\n"
	print_string += "dataset: [\n"
	print_string += "    {\n"
	print_string += "      // prettier-ignore\n"
	print_string += "      source: \n"
	jsonStr = json.dumps(statBoxPlot_data)
	print_string += jsonStr+'\n'

	chartText_2 ="""    },
    {
      transform: {
        type: 'boxplot',
        config: {
          itemNameFormatter: function (params) {
            return names[params.value];
          }
        }
      },
    },
    {
      fromDatasetIndex: 1,
      fromTransformResult: 1
    }
  ],
  dataZoom: [{id: 'dataZoomX', type: 'slider', xAxisIndex: [0], left: 10, height: 10, filterMode: 'empty', start: 0, end: 100},{id: 'dataZoomY', type: 'slider', yAxisIndex: [0], filterMode: 'empty', start: 0, end: 100}],
  tooltip: {trigger: 'item'},
  grid: {left: '20%', right: '10%', bottom: '15%'},
  yAxis: {type: 'category', boundaryGap: true, nameGap: 30, splitArea: {show: true}, splitLine: {show: true}},
  xAxis: {type: 'value', name: 'Sec', splitArea: {show: true}},
  series: [
    {
      name: 'boxplot',
      type: 'boxplot',
      datasetIndex: 1,
      tooltip: {trigger: 'item',
          formatter: function (params) {
            console.log(params.value);
            //Low = params.value[1]
          return `<u><b>${params.value[0]}</b></u>
    <table>
      <tr>
      	<td align="right">&#x2022;</td>
        <td align="left">Low   :</td>
        <td style="color:blue;"align="right"><b>${params.value[1].toFixed(2)}</b></td>
      </tr>
      <tr>
      	<td align="right">&#x2022;</td>
        <td align="left">Q1    :</td>
        <td style="color:blue;"align="right"><b>${params.value[2].toFixed(2)}</b></td>
      </tr>
      <tr>
      	<td align="right">&#x2022;</td>
        <td align="left">Q2    :</td>
        <td style="color:blue;"align="right"><b>${params.value[3].toFixed(2)}</b></td>
      </tr>
      <tr>
      	<td align="right">&#x2022;</td>
        <td align="left">Q3    :</td>
        <td style="color:blue;"align="right"><b>${params.value[4].toFixed(2)}</b></td>
      </tr>
      <tr>
      	<td align="right">&#x2022;</td>
        <td align="left">High  :</td>
        <td style="color:blue;"align="right"><b>${params.value[5].toFixed(2)}</b></td>
      </tr>  
    </table>`;              
        },    
        axisPointer: {type: 'shadow'}},      
      itemStyle: {
        borderColor: function (seriesIndex) {
          let myIndex = names.indexOf(seriesIndex.name);
          return ProfessionColor[professions[myIndex]];
                },
        borderWidth: 2
      },
      encode:{tooltip: [ 1, 2, 3, 4, 5]},
      },
    {
      name: 'outlier',
      type: 'scatter',
      encode: { x: 1, y: 0 },
      datasetIndex: 2,
    }
  ]
};
"""
	print_string += chartText_2


	myprint(chart_Output, print_string)

	chart_Output.close()
# 	end write TW5 Chart tids

def write_DPSStats_bubble_charts(uptime_Table, DPSStats, myDate, input_directory):
	#write Bubble chart tid for DPSStats
	max_fightTime = 0
	for squadDps_prof_name in uptime_Table:
		max_fightTime = max(uptime_Table[squadDps_prof_name]['duration'], max_fightTime)

	fileDate = myDate
	bubblefileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_DPSStats_TW5_Bubble_Chart.tid"
	DPSStats_bubble_chart_Output = open(bubblefileTid, "w",encoding="utf-8")
	minStatSec= 1000
	maxStatSec = 0
	
	print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
	print_string +="\ncreator: Drevarr\n"
	print_string +="tags: ChartData\n"
	print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_DPSStats_BubbleChartData\n'
	print_string +="type: application/javascript\n\n\n"

	print_string +='\nvar option = {\n  dataset: [{\n    source: ['
	print_string += '\n            ["Name", "Profession", "DPS", "Ch2DPS", "Ch5DPS", "CaDPS", "CDPS", "Downs", "Kills", "color", "Fight Time"],'
	for DPSStats_prof_name in DPSStats:
		name = DPSStats[DPSStats_prof_name]['name']
		prof = DPSStats[DPSStats_prof_name]['profession']
		fightTime = DPSStats[DPSStats_prof_name]['duration']
		myDPS = round(DPSStats[DPSStats_prof_name]['Damage_Total'] / fightTime)
		Ch2DPS = round(DPSStats[DPSStats_prof_name]['Chunk_Damage'][2] / fightTime)
		Ch5DPS = round(DPSStats[DPSStats_prof_name]['Chunk_Damage'][5] / fightTime)
		CaDPS = round(DPSStats[DPSStats_prof_name]['Carrion_Damage'] / fightTime)
		myCDPS = round(DPSStats[DPSStats_prof_name]['Coordination_Damage'] / fightTime)
		Downs = round(DPSStats[DPSStats_prof_name]['Downs'] / (fightTime / 60), 2)
		Kills = round(DPSStats[DPSStats_prof_name]['Kills'] / (fightTime / 60), 2)
		color = ProfessionColor[prof]
		if myCDPS > maxStatSec:
			maxStatSec = myCDPS
		if myCDPS < minStatSec:
			minStatSec = myCDPS
		if DPSStats[DPSStats_prof_name]['Damage_Total'] / fightTime < 500 or fightTime * 10 < max_fightTime:
			continue
		else:
			print_string += '\n            ["'+name+'", "'+prof+'", '+str(myDPS)+', '+str(Ch2DPS)+', '+str(Ch5DPS)+', '+str(CaDPS)+', '+str(myCDPS)+', '+str(Downs)+', '+str(Kills)+', "'+color+'", '+str(fightTime)+'],'
				
	print_string += '\n   ]'
	print_string += '\n  }],'
	print_string += '\n  visualMap: {'
	print_string += '\n    show: true,'
	print_string += '\n    dimension: 6, // means the 7th column		'
	print_string +='\n    min: '+str(minStatSec)+', // lower bound'
	print_string +='\n    max: '+str(maxStatSec)+', // upper bound'
	print_string +='\n    inRange: {'
	print_string +='\n      // Size of the bubble.'
	print_string +='\n      symbolSize: [5, 50]'
	print_string +='\n    }'
	print_string +='\n  },			'
	print_string +=  '\nxAxis: {'
	print_string +="\n    type: 'value',"
	print_string +='\n    name: "Ch5DPS"'
	print_string +='\n  },'
	print_string +='\n  yAxis: {'
	print_string +="\n    type: 'value',"
	print_string +='\n    name: "CaDPS"'
	print_string +='\n  },'
	print_string +="\n  tooltip: {trigger: 'axis'},"
	print_string +='\n  series: ['
	print_string +='\n    {'
	print_string +="\n      type: 'scatter',"
	print_string +='\n      encode: {'
	print_string +='\n        // Map "amount" column to x-axis.'
	print_string +="\n        x: 'Ch5DPS',"
	print_string +='\n        // Map "product" row to y-axis.'
	print_string +="\n        y: 'CaDPS',"
	print_string +='\n        // format tooltip'
	print_string +='\n        tooltip: [0, 1, 2, 3, 4, 5, 6, 7, 8, 10],'
	print_string +='\n      },	'
	print_string +='\n      itemStyle: {'
	print_string +='\n        color: function(seriesIndex) {'
	print_string +='\n        	if (seriesIndex.data[9]){'
	print_string +='\n        	  return seriesIndex.data[9];'
	print_string +='\n        	}'
	print_string +='\n        }'
	print_string +='\n      }'
	print_string +='\n    }'
	print_string +='\n  ]'
	print_string +='\n};'
		
	myprint(DPSStats_bubble_chart_Output, print_string)

	DPSStats_bubble_chart_Output.close()
#	end write bubble charts

#JEL - write TW5 Bubble Chart tids
def write_bubble_charts(players, top_players, squad_Control, myDate, input_directory):
	get_Stats = ['deaths', 'kills', 'downs', 'dmg_taken', 'dmg', 'rips', 'cleanses', 'heal', 'dist']
	boon_List = ['stability', 'protection', 'aegis', 'might', 'fury', 'resistance', 'resolution', 'quickness', 'swiftness', 'alacrity', 'vigor', 'regeneration', 'fireOut', 'shockingOut', 'frostOut', 'magneticOut', 'lightOut']
	Charts = ['kills', 'cleanse', 'rips', 'deaths', 'fury_might']
	Bubble_Chart = {}

	#for i in range(len(top_players)):
	for i in range(len(players)):
		#player = players[top_players[i]]
		player = players[i]
		prof_name = "{{"+player.profession+"}} "+player.name
		if prof_name not in Bubble_Chart:
			Bubble_Chart[prof_name]={}
			Bubble_Chart[prof_name]['name'] = player.name
			Bubble_Chart[prof_name]['profession'] = player.profession
			Bubble_Chart[prof_name]['control']=0
			Bubble_Chart[prof_name]['rips']=0
			Bubble_Chart[prof_name]['dmg']=0
			Bubble_Chart[prof_name]['cleanses']=0
			Bubble_Chart[prof_name]['heal']=0
			Bubble_Chart[prof_name]['boonScore']=0
			Bubble_Chart[prof_name]['kills']=0
			Bubble_Chart[prof_name]['downs']=0
			Bubble_Chart[prof_name]['deaths']=0
			Bubble_Chart[prof_name]['dmg_taken']=0
			Bubble_Chart[prof_name]['dist']=0
			Bubble_Chart[prof_name]['Fury_Uptime']=0
			Bubble_Chart[prof_name]['Might_Uptime']=0
		
		#gather control score per player
		sum_Control = 0
		for effect in squad_Control:
			if player.name not in squad_Control[effect]:
				continue
			else:
				sum_Control += squad_Control[effect][player.name]
		Bubble_Chart[prof_name]['control'] = sum_Control
		
		#gather boon score per player
		sum_Boons = 0
		for boon in boon_List:
			sum_Boons += player.average_stats[boon]
		#for aura in auras_TableOut:
		#	if player.name in auras_TableOut[aura]:
		#		sum_Boons += (auras_TableOut[aura][player.name] / player.duration_fights_present)
		Bubble_Chart[prof_name]['boonScore'] = round(sum_Boons, 4)
		
		#gather Stats scores per player
		for statItem in get_Stats:
			Bubble_Chart[prof_name][statItem] = player.average_stats[statItem]

		#Calculate Fury Uptime per player
		if prof_name in uptime_Table and 'fury' in uptime_Table[prof_name] and uptime_Table[prof_name]['duration'] >0:
			Bubble_Chart[prof_name]['Fury_Uptime'] = round((uptime_Table[prof_name]['fury']/uptime_Table[prof_name]['duration'])*100,2)
		else:
			Bubble_Chart[prof_name]['Fury_Uptime'] = 0.00

		#Calculate Avg_Might per player
		if prof_name in uptime_Table and 'might' in uptime_Table[prof_name] and uptime_Table[prof_name]['duration'] >0:
			Bubble_Chart[prof_name]['Might_Uptime'] = round((uptime_Table[prof_name]['might']/uptime_Table[prof_name]['duration'])*100,2)
		else:
			Bubble_Chart[prof_name]['Might_Uptime'] = 0.00
			
	for chart in Charts:
		fileDate = myDate
		bubblefileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_"+chart+"_TW5_Bubble_Chart.tid"
		bubble_chart_Output = open(bubblefileTid, "w",encoding="utf-8")
		minStatSec= 1000
		maxStatSec = 0
		
		print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
		print_string +="\ncreator: Drevarr\n"
		print_string +="tags: ChartData\n"
		print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_'+chart+'_BubbleChartData\n'
		print_string +="type: application/javascript\n\n\n"

		print_string +='\nvar option = {\n  dataset: [{\n    source: ['
		
		if chart == 'kills':
			print_string += '\n            ["Name", "Profession", "Kills", "Downs", "DPS", "color"],'
			for prof_name in Bubble_Chart:
				color = ProfessionColor[Bubble_Chart[prof_name]['profession']]
				print_string += '\n            ["'+Bubble_Chart[prof_name]['name']+'", "'+Bubble_Chart[prof_name]['profession']+'", '+str(Bubble_Chart[prof_name]['kills'])+', '+str(Bubble_Chart[prof_name]['downs'])+', '+str(Bubble_Chart[prof_name]['dmg'])+', "'+color+'"],'
				if Bubble_Chart[prof_name]['dmg'] > maxStatSec:
					maxStatSec = Bubble_Chart[prof_name]['dmg']
				if Bubble_Chart[prof_name]['dmg'] < minStatSec:
					minStatSec = Bubble_Chart[prof_name]['dmg']

		if chart == 'cleanse':
			print_string += '\n            ["Name", "Profession", "Cleanses", "Heals", "Boon Score", "color"],'
			for prof_name in Bubble_Chart:
				color = ProfessionColor[Bubble_Chart[prof_name]['profession']]
				print_string += '\n            ["'+Bubble_Chart[prof_name]['name']+'", "'+Bubble_Chart[prof_name]['profession']+'", '+str(Bubble_Chart[prof_name]['cleanses'])+', '+str(Bubble_Chart[prof_name]['heal'])+', '+str(Bubble_Chart[prof_name]['boonScore'])+', "'+color+'"],'
				if Bubble_Chart[prof_name]['boonScore'] > maxStatSec:
					maxStatSec = Bubble_Chart[prof_name]['boonScore']
				if Bubble_Chart[prof_name]['boonScore'] < minStatSec:
					minStatSec = Bubble_Chart[prof_name]['boonScore']
				
		if chart == 'rips':
			print_string += '\n            ["Name", "Profession", "Strips", "Control", "DPS", "color"],'
			for prof_name in Bubble_Chart:
				color = ProfessionColor[Bubble_Chart[prof_name]['profession']]
				print_string += '\n            ["'+Bubble_Chart[prof_name]['name']+'", "'+Bubble_Chart[prof_name]['profession']+'", '+str(Bubble_Chart[prof_name]['rips'])+', '+str(Bubble_Chart[prof_name]['control'])+', '+str(Bubble_Chart[prof_name]['dmg'])+', "'+color+'"],'		
				if Bubble_Chart[prof_name]['dmg'] > maxStatSec:
					maxStatSec = Bubble_Chart[prof_name]['dmg']
				if Bubble_Chart[prof_name]['dmg'] < minStatSec:
					minStatSec = Bubble_Chart[prof_name]['dmg']
				
		if chart == 'deaths':
			print_string += '\n            ["Name", "Profession", "Deaths", "Damage_Taken", "Distance_to_Tag", "color"],'
			for prof_name in Bubble_Chart:
				color = ProfessionColor[Bubble_Chart[prof_name]['profession']]
				print_string += '\n            ["'+Bubble_Chart[prof_name]['name']+'", "'+Bubble_Chart[prof_name]['profession']+'", '+str(Bubble_Chart[prof_name]['deaths'])+', '+str(Bubble_Chart[prof_name]['dmg_taken'])+', '+str(Bubble_Chart[prof_name]['dist'])+', "'+color+'"],'
				if Bubble_Chart[prof_name]['dist'] > maxStatSec:
					maxStatSec = Bubble_Chart[prof_name]['dist']
				if Bubble_Chart[prof_name]['dist'] < minStatSec:
					minStatSec = Bubble_Chart[prof_name]['dist']

		if chart == 'fury_might':
			print_string += '\n            ["Name", "Profession", "Fury", "Might", "DPS", "color"],'
			for prof_name in Bubble_Chart:
				color = ProfessionColor[Bubble_Chart[prof_name]['profession']]
				print_string += '\n            ["'+Bubble_Chart[prof_name]['name']+'", "'+Bubble_Chart[prof_name]['profession']+'", '+str(Bubble_Chart[prof_name]['Fury_Uptime'])+', '+str(Bubble_Chart[prof_name]['Might_Uptime'])+', '+str(Bubble_Chart[prof_name]['dmg'])+', "'+color+'"],'
				if Bubble_Chart[prof_name]['dmg'] > maxStatSec:
					maxStatSec = Bubble_Chart[prof_name]['dmg']
				if Bubble_Chart[prof_name]['dmg'] < minStatSec:
					minStatSec = Bubble_Chart[prof_name]['dmg']

		print_string += '\n   ]'
		print_string += '\n  }],'
		print_string += '\n  visualMap: {'
		print_string += '\n    show: true,'
		print_string += '\n    dimension: 4, // means the 5th column		'
		print_string +='\n    min: '+str(minStatSec)+', // lower bound'
		print_string +='\n    max: '+str(maxStatSec)+', // upper bound'
		print_string +='\n    inRange: {'
		print_string +='\n      // Size of the bubble.'
		print_string +='\n      symbolSize: [5, 50]'
		print_string +='\n    }'
		print_string +='\n  },			'
		print_string +=  '\nxAxis: {'
		print_string +="\n    type: 'value',"
		if chart == 'kills':
			print_string +='\n    name: "Downs per Second"'
		if chart == 'cleanse':
			print_string +='\n    name: "Cleanses per Second"'	
		if chart == 'deaths':
			print_string +='\n    name: "Average Deaths"'		
		if chart == 'rips':
			print_string +='\n    name: "Control Effect Score"'			
		if chart == 'fury_might':
			print_string +='\n    name: "Fury Uptime"'			
		print_string +='\n  },'
		print_string +='\n  yAxis: {'
		print_string +="\n    type: 'value',"
		if chart == 'kills':
			print_string +='\n    name: "Kills per Second"'
		if chart == 'cleanse':
			print_string +='\n    name: "Heals per Second"'	
		if chart == 'deaths':
			print_string +='\n    name: "Average Damage Taken"'		
		if chart == 'rips':
			print_string +='\n    name: "Strips per Second"'
		if chart == 'fury_might':
			print_string +='\n    name: "Might Uptime"'						
		print_string +='\n  },'
		print_string +="\n  tooltip: {trigger: 'axis',\n        axisPointer: {\n          type: 'cross'\n        },    \n},"
		print_string +='\n  series: ['
		print_string +='\n    {'
		print_string +="\n      type: 'scatter',"
		print_string +='\n      encode: {'
		print_string +='\n        // Map "amount" column to x-axis.'
		if chart == 'kills':
			print_string +="\n        x: 'Downs',"
		if chart == 'cleanse':
			print_string +="\n        x: 'Cleanses',"		
		if chart == 'deaths':
			print_string +="\n        x: 'Deaths',"
		if chart == 'rips':
			print_string +="\n        x: 'Control',"	
		if chart == 'fury_might':
			print_string +="\n        x: 'Fury',"				
		print_string +='\n        // Map "product" row to y-axis.'
		if chart == 'kills':	
			print_string +="\n        y: 'Kills',"
		if chart == 'cleanse':	
			print_string +="\n        y: 'Heals',"
		if chart == 'deaths':	
			print_string +="\n        y: 'Damage_Taken',"
		if chart == 'rips':	
			print_string +="\n        y: 'Strips',"	
		if chart == 'fury_might':	
			print_string +="\n        y: 'Might',"				
		print_string +='\n        // format tooltip'
		print_string +='\n        tooltip: [0, 1, 2, 3, 4],'
		print_string +='\n      },	'
		print_string +='\n      itemStyle: {'
		print_string +='\n        color: function(seriesIndex) {'
		print_string +='\n          console.log(seriesIndex);'
		print_string +='\n        	console.log(seriesIndex.color);'
		print_string +='\n        	console.log(seriesIndex.data[5]);'
		print_string +='\n        	if (seriesIndex.data[5]){'
		print_string +='\n        	  return seriesIndex.data[5];'
		print_string +='\n        	}'
		print_string +='\n        }'
		print_string +='\n      }'
		print_string +='\n    }'
		print_string +='\n  ]'
		print_string +='\n};'
		
		myprint(bubble_chart_Output, print_string)

		bubble_chart_Output.close()
#	end write bubble charts

def write_box_plot_charts(DPS_List, myDate, input_directory, ChartType):
	Charts = ['Profession', 'Profession_and_Name']
	fileDate = myDate
	for chart in Charts:
		boxPlotfileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_"+ChartType+"_"+chart+"_TW5_Box_Plot_Chart.tid"
		boxPlot_chart_Output = open(boxPlotfileTid, "w",encoding="utf-8")

		print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
		print_string +="\ncreator: Drevarr\n"
		print_string +="tags: ChartData\n"
		print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_'+ChartType+'_'+chart+'_Box_PlotChartData\n'
		print_string +="type: application/javascript\n"

		#print_string +='const colors = '
		print_string +='\nconst professions = '
		if chart == 'Profession':
			sorted_DPS_List = OrderedDict(sorted(DPS_List['prof'].items()))
			print_string += str(list(sorted_DPS_List.keys()))
		if chart == 'Profession_and_Name':
			sorted_DPS_List = OrderedDict(sorted(DPS_List['prof_name'].items()))
			print_string += str(list(sorted_DPS_List.keys()))
		print_string +='\n'
		print_string +='\nProfessionColor = {"Warrior":"#FFD166", "Berserker":"#B39247", "Spellbreaker":"#665429", "Bladesworm":"#19150A", "Guardian":"#72C1D9", "Dragonhunter":"#508798", "Firebrand":"#2E4D57", "Willbender":"#0B1316", "Revenant":"#D16E5A", "Herald":"#924D3F", "Renegade":"#542C24", "Vindicator":"#2A1612", "Engineer":"#D09C59", "Scrapper":"#926D3E", "Holosmith":"#533E24", "Mechanist":"#2A1F12", "Ranger":"#8CDC82", "Druid":"#629A5B", "Soulbeast":"#385834", "Untamed":"#1C2C1A", "Thief":"#C08F95", "Daredevil":"#866468", "Deadeye":"#4D393C", "Specter":"#261D1E", "Elementalist":"#F68A87", "Tempest":"#AC615F", "Weaver":"#623736", "Catalyst":"#311C1B", "Mesmer":"#B679D5", "Chronomancer":"#7F5595", "Mirage":"#493055", "Virtuoso":"#24182B", "Necromancer":"#52A76F", "Reaper":"#39754E", "Scourge":"#21432C", "Harbinger":"#08110B"}'
		print_string +='\noption = {'
		print_string +='\n  title: ['
		print_string +="\n    {text: '"+ChartType+" by "+chart+" across all fights', left: 'center'},"
		print_string +="\n    {text: '"+ChartType+" across all fights \\nupper: Q3 + 1.5 * IQR \\nlower: Q1 - 1.5 * IQR', borderColor: '#999', borderWidth: 1, textStyle: {fontSize: 10}, left: '1%', top: '90%'}"
		print_string +="\n  ],"
		print_string +="\ndataset: ["
		print_string +="\n    {"
		print_string +="\n      // prettier-ignore"
		print_string +="\n      source: "
		if chart == 'Profession':
			sorted_DPS_List = OrderedDict(sorted(DPS_List['prof'].items()))
			print_string += str(list(sorted_DPS_List.values()))
		if chart == 'Profession_and_Name':
			sorted_DPS_List = OrderedDict(sorted(DPS_List['prof_name'].items()))
			print_string += str(list(sorted_DPS_List.values()))
		print_string += "\n    },"
		print_string += "\n    {"
		print_string += "\n      transform: {"
		print_string += "\n        type: 'boxplot',"
		print_string += "\n        config: {"
		print_string += "\n          itemNameFormatter: function (params) {"
		print_string += "\n            return professions[params.value];"
		print_string += "\n          }"
		print_string += "\n        }"
		print_string += "\n      },"
		print_string += "\n    },"
		print_string += "\n    {"
		print_string += "\n      fromDatasetIndex: 1,"
		print_string += "\n      fromTransformResult: 1"
		print_string += "\n    }"
		print_string += "\n  ],"
		if chart == 'Profession':
			print_string += "\n  dataZoom: [{id: 'dataZoomX', type: 'slider', xAxisIndex: [0], left: 10, height: 10, filterMode: 'empty', start: 0, end: 100},{id: 'dataZoomY', type: 'slider', yAxisIndex: [0], filterMode: 'empty', start: 0, end: 100}],"
		if chart == 'Profession_and_Name':
			print_string += "\n  dataZoom: [{id: 'dataZoomX', type: 'slider', xAxisIndex: [0], left: 10, height: 10, filterMode: 'empty', start: 0, end: 100},{id: 'dataZoomY', type: 'slider', yAxisIndex: [0], filterMode: 'empty', start: 0, end: 30}],"
		print_string += "\n  tooltip: {trigger: 'item', axisPointer: {type: 'shadow'}},"
		print_string += "\n  grid: {left: '10%', right: '10%', bottom: '15%'},"
		print_string += "\n  yAxis: {type: 'category', boundaryGap: true, nameGap: 30, splitArea: {show: true}, splitLine: {show: true}},"
		print_string += "\n  xAxis: {type: 'value', name: '"+ChartType+"', splitArea: {show: true}},"
		print_string += "\n  series: ["
		print_string += "\n    {"
		print_string += "\n      name: 'boxplot',"
		print_string += "\n      type: 'boxplot',"
		print_string += "\n      datasetIndex: 1,"
		print_string += "\n      itemStyle: {"
		print_string += "\n        borderColor: function (seriesIndex) {  "
		print_string += "\n          return ProfessionColor[seriesIndex.name.split('_', 1)];"
		print_string += "\n                }"
		print_string += "\n      },"
		print_string += "\n      encode:{tooltip: [ 1, 2, 3, 4, 5]},"
		print_string += "\n      },\n    {"
		print_string += "\n      name: 'outlier',"
		print_string += "\n      type: 'scatter',"
		print_string += "\n      encode: { x: 1, y: 0 },"
		print_string += "\n      datasetIndex: 2,"
		print_string += "\n      itemStyle: {"
		print_string += "\n        color: function (seriesIndex) {  "
		print_string += "\n          return ProfessionColor[seriesIndex.name.split('_', 1)];"
		print_string += "\n                }"
		print_string += "\n      },"		
		print_string += "\n    }\n  ]\n};		"

		
		myprint(boxPlot_chart_Output, print_string)

		boxPlot_chart_Output.close()
#	end write bubble charts

#Start Heat Map Chart
def write_spike_damage_heatmap(squad_damage_output, myDate, input_directory):
	fileDate = myDate
	fileTid = input_directory+"/"+fileDate.strftime('%Y%m%d%H%M')+"_spike_damage_heatmap_TW5_Chart.tid"
	heatmap_Output = open(fileTid, "w",encoding="utf-8")
	fight_counter = 0
	fight_output = []
	fight_list = []
	fight_seconds = [i for i in range(1,61)]
	for fight in squad_damage_output:
		fight_list.append(fight)
		if squad_damage_output[fight].values():
			max_squad_dmg = max(squad_damage_output[fight].values())
		else:
			max_squad_dmg = 1
		#print(fight_header)
		for i in range(1, 61):
			if i in squad_damage_output[fight]:
				squad_dmg = squad_damage_output[fight][i]
			else:
				squad_dmg = 0
			fight_sec_data = [fight_counter, i, round((squad_dmg/max_squad_dmg), 1)]
			fight_output.append(fight_sec_data)
		fight_counter +=1

	print_string = 'created: '+fileDate.strftime("%Y%m%d%H%M%S")
	print_string +="\ncreator: Drevarr\n"
	print_string +="tags: ChartData\n"
	print_string +='title: '+fileDate.strftime("%Y%m%d%H%M")+'_spike_damage_heatmap_ChartData\n'
	print_string +="type: application/javascript\n\n\n"
	print_string +="// prettier-ignore\n"
	#output HeatMap Seconds
	jsonStr = json.dumps(fight_seconds)
	print_string +='const seconds = '+jsonStr+';\n'
	#output HeatMap Fights
	jsonStr = json.dumps(fight_list)
	print_string +='const fights = '+jsonStr+';\n'
	#output data
	jsonStr = json.dumps(fight_output)
	print_string +='const data = '+jsonStr+'\n'
	print_string +="    .map(function (item) {\n"
	print_string +="    return [item[1], item[0], item[2] || '-'];\n"
	print_string +="});	\n"
	print_string += "option = {\n"
	print_string += "  title: {\n"
	print_string += "    left: 'center',\n"
	print_string += "    text: 'Damage  / Max Damage in 1 Second\\n(limited to first 60 seconds of fight)'\n"
	print_string +="},\n"
	heatMapText ="""  tooltip: {
    position: 'top',
  },
  grid: {
    height: '80%',
    left: '15%',
    top: '10%'
  },
  xAxis: {
    type: 'category',
    data: seconds,
    splitArea: {
      show: true
    }
  },
  yAxis: {
    type: 'category',
    data: fights,
    name: 'Fight Ending',
    splitArea: {
      show: true
    }
  },
  visualMap: {
    min: 0,
    max: 1,
    calculable: true,
    orient: 'vertical',
    left: 'left',
    bottom: '55%'
  },
  dataZoom: [
    {
      type: 'slider',
      show: true,
      xAxisIndex: [0],
      start: 0,
      end: 30
    },
    {
      type: 'inside',
      xAxisIndex: [0],
      start: 0,
      end: 30
    },
    {
      type: 'slider',
      show: true,
      yAxisIndex: [0],
      start: 0,
      end: 30
    },
    {
      type: 'inside',
      yAxisIndex: [0],
      start: 0,
      end: 30
    },    
  ],
  series: [
    {
      name: 'Spike Damage',
      type: 'heatmap',
      data: data,
      label: {
        show: true
      },
      emphasis: {
        itemStyle: {
          shadowBlur: 25,
          shadowColor: 'rgba(0, 0, 0, 0.5)'
        }
      }
    }
  ]
};
"""
	print_string += heatMapText


	myprint(heatmap_Output, print_string)

	heatmap_Output.close()
#end Heat Map Chart	

def write_to_json(overall_raid_stats, overall_squad_stats, fights, players, top_total_stat_players, top_average_stat_players, top_consistent_stat_players, top_percentage_stat_players, top_late_players, top_jack_of_all_trades_players, squad_offensive, squad_Control, enemy_Control, enemy_Control_Player, downed_Healing, uptime_Table, stacking_uptime_Table, auras_TableOut, Death_OnTag, Attendance, DPS_List, CPS_List, SPS_List, HPS_List, DPSStats, output_file):
	json_dict = {}
	json_dict["overall_raid_stats"] = {key: value for key, value in overall_raid_stats.items()}
	json_dict["overall_squad_stats"] = {key: value for key, value in overall_squad_stats.items()}
	json_dict["fights"] = [jsons.dump(fight) for fight in fights]
	json_dict["players"] = [jsons.dump(player) for player in players]
	json_dict["top_total_players"] =  {key: value for key, value in top_total_stat_players.items()}
	json_dict["top_average_players"] =  {key: value for key, value in top_average_stat_players.items()}
	json_dict["top_consistent_players"] =  {key: value for key, value in top_consistent_stat_players.items()}
	json_dict["top_percentage_players"] =  {key: value for key, value in top_percentage_stat_players.items()}
	json_dict["top_late_players"] =  {key: value for key, value in top_late_players.items()}
	json_dict["top_jack_of_all_trades_players"] =  {key: value for key, value in top_jack_of_all_trades_players.items()}
	json_dict["squad_offensive"] =  {key: value for key, value in squad_offensive.items()}
	json_dict["squad_Control"] =  {key: value for key, value in squad_Control.items()}
	json_dict["enemy_Control"] =  {key: value for key, value in enemy_Control.items()}
	json_dict["enemy_Control_Player"] =  {key: value for key, value in enemy_Control_Player.items()}
	json_dict["uptime_Table"] =  {key: value for key, value in uptime_Table.items()}
	json_dict["stacking_uptime_Table"] =  {key: value for key, value in stacking_uptime_Table.items()}
	json_dict["auras_TableOut"] =  {key: value for key, value in auras_TableOut.items()}
	json_dict["Death_OnTag"] =  {key: value for key, value in Death_OnTag.items()}
	json_dict["Attendance"] =  {key: value for key, value in Attendance.items()}
	json_dict["DPS_List"] =  {key: value for key, value in DPS_List.items()}
	json_dict["CPS_List"] =  {key: value for key, value in CPS_List.items()}
	json_dict["SPS_List"] =  {key: value for key, value in SPS_List.items()}
	json_dict["HPS_List"] =  {key: value for key, value in HPS_List.items()}
	json_dict["DPSStats"] =  {key: value for key, value in DPSStats.items()}
	json_dict["downed_Healing"] =  {key: value for key, value in downed_Healing.items()}
	json_dict["MOA_Targets"] =  {key: value for key, value in MOA_Targets.items()}
	json_dict["MOA_Casters"] =  {key: value for key, value in MOA_Casters.items()}
	json_dict["Buffs_Personal"] =  {key: value for key, value in buffs_personal.items()}
	json_dict["squad_damage_output"] =  {key: value for key, value in squad_damage_output.items()}
	json_dict["skill_Dict"] =  {key: value for key, value in skill_Dict.items()}

		
	with open(output_file, 'w') as json_file:
		json.dump(json_dict, json_file, indent=4)

def calc_weighted_dps_enemy(players, fights):
	for player in players:
		playerDamages = []
		playerDurations = []
		playerEnemies = []
		weighted_DPS_Enemy = []
		sum_playerDurations = 0

		for fight in player['stats_per_fight']:
			playerDamages.append(fight['dmg'])
			playerDurations.append(fight['time_in_combat'])
		for fight in fights:
			playerEnemies.append(fight['enemies'])

		for fightTime in playerDurations:
			if fightTime >0:
				sum_playerDurations += fightTime

		for (dmg, enemy, duration) in zip(playerDamages, playerEnemies, playerDurations):
			if dmg != -1:
				weighted_DPS_Enemy.append(round((dmg/enemy) * (duration / sum_playerDurations),2))

	return sum(weighted_DPS_Enemy)