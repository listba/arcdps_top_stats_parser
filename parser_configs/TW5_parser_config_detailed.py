#    This file contains the configuration for computing the detailed top stats in arcdps logs as parsed by Elite Insights.
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

stats_to_compute = ['deaths', 'downed', 'iol', 'res', 'kills', 'downs','dmg', 'Pdmg', 'Cdmg', 'dmg_taken', 'rips', 'cleanses', 'superspeed', 'stealth', 'HiS', 'dist', 'stability', 'protection', 'aegis', 'might', 'fury', 'resistance', 'resolution', 'quickness', 'swiftness', 'alacrity', 'vigor', 'regeneration', 'heal', 'barrier', 'barrierDamage', 'swaps', 'dodges', 'evades', 'invulns', 'hitsMissed', 'interupted', 'blocks', 'fireOut', 'shockingOut', 'frostOut', 'magneticOut', 'lightOut', 'darkOut', 'chaosOut', 'ripsIn', 'cleansesIn', 'downContrib']
aurasIn_to_compute = []
aurasOut_to_compute = ['fireOut', 'shockingOut', 'frostOut', 'magneticOut', 'lightOut', 'darkOut', 'chaosOut']
defenses_to_compute = ['dmg_taken', 'barrierDamage', 'hitsMissed', 'interupted', 'invulns', 'evades', 'blocks', 'dodges', 'cleansesIn', 'ripsIn', 'downed', 'deaths']
#defense_to_compute =['dmg_taken','blockedCount', 'evadedCount', 'missedCount', 'dodgeCount', 'invulnedCount', 'damageBarrier', 'interruptedCount', 'downCount', 'deadCount']

# How many players will be listed who achieved top stats most often for each stat?
num_players_listed = {'dmg': 50, 'Pdmg': 50, 'Cdmg': 50, 'iol': 50,'rips': 50, 'cleanses': 50, 'dist': 50, 'stability': 50, 'protection': 50, 'aegis': 50, 'might': 50, 'fury': 50, 'dmg_taken': 50, 'deaths': 50, 'downed': 50,'res': 50, 'superspeed': 50, 'stealth': 50, 'HiS': 50, 'resistance': 50, 'resolution': 50, 'quickness': 50, 'swiftness': 50, 'alacrity': 50, 'vigor': 50, 'regeneration': 50, 'heal': 50, 'barrier': 50, 'barrierDamage': 50, 'swaps': 50, 'kills': 50, 'downs': 50, 'dodges': 50, 'evades': 50, 'hitsMissed': 50, 'interupted': 50, 'blocks': 50, 'invulns': 50, 'fireOut': 50, 'shockingOut': 50, 'frostOut': 50, 'magneticOut': 50, 'lightOut': 50, 'darkOut': 50, 'chaosOut': 50, 'ripsIn': 50, 'cleansesIn': 50, 'downContrib': 50}
# What portion (%) of are considered to be "top" in each fight for each stat?
num_players_considered_top_percentage = 5

# For what portion of all fights does a player need to be there to be considered for "consistency percentage" awards?
attendance_percentage_for_percentage = 50
# For what portion of all fights does a player need to be there to be considered for "late but great" awards?
attendance_percentage_for_late = 50
# For what portion of all fights does a player need to be there to be considered for "jack of all trades" awards? 
attendance_percentage_for_buildswap = 30
# For what portion of all fights does a player need to be there to be considered for "top average" awards? 
attendance_percentage_for_average = 33

# What portion of the top total player stat does someone need to reach to be considered for total awards?
percentage_of_top_for_consistent = 10
# What portion of the total stat of the top consistent player does someone need to reach to be considered for consistency awards?
percentage_of_top_for_total = 10
# What portion of the percentage the top consistent player reached top does someone need to reach to be considered for percentage awards?
percentage_of_top_for_percentage = 10
# What portion of the percentage the top consistent player reached top does someone need to reach to be considered for late but great awards?
percentage_of_top_for_late = 75
# What portion of the percentage the top consistent player reached top does someone need to reach to be considered for jack of all trades awards?
percentage_of_top_for_buildswap = 75

# minimum number of allied players to consider a fight in the stats
min_allied_players = 5
# minimum duration of a fight to be considered in the stats
min_fight_duration = 20
# minimum number of enemies to consider a fight in the stats
min_enemy_players = 5

# Produce Charts for stats_to_compute
charts = True
# Include the Squad Comp and Fight Review tabs
include_comp_and_review = True

# names as which each specialization will show up in the stats
profession_abbreviations = {}
profession_abbreviations["Guardian"] = "Guardian"
profession_abbreviations["Dragonhunter"] = "Dragonhunter"
profession_abbreviations["Firebrand"] = "Firebrand"
profession_abbreviations["Willbender"] = "Willbender"

profession_abbreviations["Revenant"] = "Revenant"
profession_abbreviations["Herald"] = "Herald"
profession_abbreviations["Renegade"] = "Renegade"
profession_abbreviations["Vindicator"] = "Vindicator"    

profession_abbreviations["Warrior"] = "Warrior"
profession_abbreviations["Berserker"] = "Berserker"
profession_abbreviations["Spellbreaker"] = "Spellbreaker"
profession_abbreviations["Bladesworn"] = "Bladesworn"

profession_abbreviations["Engineer"] = "Engineer"
profession_abbreviations["Scrapper"] = "Scrapper"
profession_abbreviations["Holosmith"] = "Holosmith"
profession_abbreviations["Mechanist"] = "Mechanist"    

profession_abbreviations["Ranger"] = "Ranger"
profession_abbreviations["Druid"] = "Druid"
profession_abbreviations["Soulbeast"] = "Soulbeast"
profession_abbreviations["Untamed"] = "Untamed"    

profession_abbreviations["Thief"] = "Thief"
profession_abbreviations["Daredevil"] = "Daredevil"
profession_abbreviations["Deadeye"] = "Deadeye"
profession_abbreviations["Specter"] = "Specter"

profession_abbreviations["Elementalist"] = "Elementalist"
profession_abbreviations["Tempest"] = "Tempest"
profession_abbreviations["Weaver"] = "Weaver"
profession_abbreviations["Catalyst"] = "Catalyst"

profession_abbreviations["Mesmer"] = "Mesmer"
profession_abbreviations["Chronomancer"] = "Chronomancer"
profession_abbreviations["Mirage"] = "Mirage"
profession_abbreviations["Virtuoso"] = "Virtuoso"
    
profession_abbreviations["Necromancer"] = "Necromancer"
profession_abbreviations["Reaper"] = "Reaper"
profession_abbreviations["Scourge"] = "Scourge"
profession_abbreviations["Harbinger"] = "Harbinger"

# name each stat will be written as
stat_names = {}
stat_names["dmg"] = "Damage"
stat_names["Pdmg"] = "Power Damage"
stat_names["Cdmg"] = "Condi Damage"
stat_names["dmg_taken"] = "Damage Taken"
stat_names["rips"] = "Boon Strips"
stat_names["stability"] = "Stability"
stat_names["protection"] = "Protection"
stat_names["aegis"] = "Aegis"
stat_names["might"] = "Might"
stat_names["fury"] = "Fury"
stat_names["cleanses"] = "Condition Cleanses"
stat_names["heal"] = "Healing"
stat_names["barrier"] = "Barrier"
stat_names["barrierDamage"] = "Barrier Damage"
stat_names["dist"] = "Distance to Tag"
stat_names["deaths"] = "Deaths"
stat_names["downed"] = "Downed"
stat_names["superspeed"] = "Superspeed"
stat_names["stealth"] = "Stealth"
stat_names["HiS"] = "Hide in Shadows"
stat_names["regeneration"] = "Regeneration"
stat_names["resistance"] = "Resistance"
stat_names["resolution"] = "Resolution"
stat_names["quickness"] = "Quickness"
stat_names["swiftness"] = "Swiftness"
stat_names["alacrity"] = "Alacrity"
stat_names["vigor"] = "Vigor"
stat_names["res"] = "Resurrect"
stat_names["iol"] = "Illusion of Life"
stat_names["cripple"] = "Cripple"
stat_names["weakness"] = "Weakness"
stat_names["daze"] = "Daze"
stat_names["immobilize"] = "Immobilize"
stat_names["swaps"] = "Weapon Swaps"
stat_names["kills"] = "Enemies Killed"
stat_names["downs"] = "Enemies Downed"
stat_names["dodges"] = "Dodge Attempts"
stat_names["evades"] = "Evaded Attacks"
stat_names["blocks"] = "Blocked Attacks"
stat_names["invulns"] = "Invulnerable to Attacks"
stat_names["interupted"] = "Interupted"
stat_names["hitsMissed"] = "Hits Missed Against"
stat_names["fireOut"] = "Fire Aura"
stat_names["shockingOut"] = "Shocking Aura"
stat_names["frostOut"] = "Frost Aura"
stat_names["magneticOut"] = "Magnetic Aura"
stat_names["lightOut"] = "Light Aura"
stat_names["darkOut"] = "Dark Aura"
stat_names["chaosOut"] = "Chaos Aura"
stat_names["ripsIn"] = "Boon Strips Incoming"
#stat_names["ripsTime"] = "Boon Time Lost"
stat_names["cleansesIn"] = "Condition Cleanses Incoming"
#stat_names["cleansesTime"] = "Condition Time Cleared"
stat_names["downContrib"] = "Down Contribution in Damage"