import xlsxwriter
import json
import requests
from os import listdir
from os.path import isfile, join


# Convert minecraft name to UUID. Requires internet access!
# Argument: string in form of name of user
# Return: string in form of uuid of player
def convert_minecraft_name_to_uuid(name):
    api_request = requests.get("https://api.mojang.com/users/profiles/minecraft/" + name)
    json_data = api_request.json()
    if "id" in json_data:
        return json_data["id"]
    else:
        return "Error. uuid of " + name + " not found."


# Convert number to column number
# Parameter: Integer in form of column number
# Return: String
def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


# Excluded UUID's
excluded_uuid_list = ["d4de507e-d284-4470-8e15-18b752284a04"]
excluded_player_names_list = ["BearVintage", "ChemistryGuy"]
for player_name in excluded_player_names_list:
    short_player_uuid = convert_minecraft_name_to_uuid(player_name)
    excluded_uuid_list.append(short_player_uuid[:8:] + '-' + short_player_uuid[8:12:] + '-' + short_player_uuid[12:16:] + '-' + short_player_uuid[16:20:] + '-' + short_player_uuid[20::])

# Colors
light_magenta = "#FF80FF"
light_pink = "#FDDCE8"

# Get name of all stats files
path = "D:\\Users\\vanro\\Documenten\\Data\\MinecraftBackups\\stats"
players = [f for f in listdir(path) if isfile(join(path, f)) and f[-4::] == "json" and f[:-5:] not in excluded_uuid_list]

# Get player_data
player_data = []
player_uuid_name_dict = {}
for player in players:
    file = open(path + "\\" + player, "r")
    name_fetch = requests.get("https://sessionserver.mojang.com/session/minecraft/profile/" + player.replace("-", "")[:-5:]).json()
    if "name" in name_fetch:
        player_uuid_name_dict[player[:-5:]] = name_fetch["name"]
    else:
        player_uuid_name_dict[player[:-5:]] = "Fetch failed. Please try again in a minute."
    player_data.append([player[:-5:], json.loads(file.readline())])


# Get first rows
killed_by_first_row = []
crafted_first_row = []
mined_first_row = []
used_first_row = []
killed_first_row = []
dropped_first_row = []
picked_up_first_row = []
custom_first_row = []
broken_first_row = []
for statistics in player_data:
    stats = statistics[1]["stats"]
    if "minecraft:killed_by" in stats:
        for stat in stats["minecraft:killed_by"]:
            if stat not in killed_by_first_row:
                killed_by_first_row.append(stat)
    if "minecraft:crafted" in stats:
        for stat in stats["minecraft:crafted"]:
            if stat not in crafted_first_row:
                crafted_first_row.append(stat)
    if "minecraft:mined" in stats:
        for stat in stats["minecraft:mined"]:
            if stat not in mined_first_row:
                mined_first_row.append(stat)
    if "minecraft:used" in stats:
        for stat in stats["minecraft:used"]:
            if stat not in used_first_row:
                used_first_row.append(stat)
    if "minecraft:killed" in stats:
        for stat in stats["minecraft:killed"]:
            if stat not in killed_first_row:
                killed_first_row.append(stat)
    if "minecraft:dropped" in stats:
        for stat in stats["minecraft:dropped"]:
            if stat not in dropped_first_row:
                dropped_first_row.append(stat)
    if "minecraft:picked_up" in stats:
        for stat in stats["minecraft:picked_up"]:
            if stat not in picked_up_first_row:
                picked_up_first_row.append(stat)
    if "minecraft:custom" in stats:
        for stat in stats["minecraft:custom"]:
            if stat not in custom_first_row:
                custom_first_row.append(stat)
    if "minecraft:broken" in stats:
        for stat in stats["minecraft:broken"]:
            if stat not in broken_first_row:
                broken_first_row.append(stat)

# Create a workbook

workbook = xlsxwriter.Workbook('statistics.xlsx')

# Formats
cell_light_magenta = workbook.add_format({"bg_color": light_magenta, "bold": True})
cell_light_pink = workbook.add_format({"bg_color": light_pink})
cell_first_row = workbook.add_format({"bold": True, "right": 2})
cell_first_column = workbook.add_format({"bold": True, "bottom": 2})
cell_stat = workbook.add_format({"right": 3, "bottom": 3})

#  Add worksheets and first rows.
worksheets = []
killed_by_worksheet = workbook.add_worksheet("Killed by")
worksheets.append(killed_by_worksheet)
for i in range(0, len(killed_by_first_row)):
    killed_by_worksheet.write(0, i + 1, killed_by_first_row[i][10::].replace("_", " ").title(), cell_first_column)
crafted_worksheet = workbook.add_worksheet("Crafted")
worksheets.append(crafted_worksheet)
for i in range(0, len(crafted_first_row)):
    crafted_worksheet.write(0, i + 1, crafted_first_row[i][10::].replace("_", " ").title(), cell_first_column)
mined_worksheet = workbook.add_worksheet("Mined")
worksheets.append(mined_worksheet)
for i in range(0, len(mined_first_row)):
    mined_worksheet.write(0, i + 1, mined_first_row[i][10::].replace("_", " ").title(), cell_first_column)
used_worksheet = workbook.add_worksheet("Used")
worksheets.append(used_worksheet)
for i in range(0, len(used_first_row)):
    used_worksheet.write(0, i + 1, used_first_row[i][10::].replace("_", " ").title(), cell_first_column)
killed_worksheet = workbook.add_worksheet("Killed")
worksheets.append(killed_worksheet)
for i in range(0, len(killed_first_row)):
    killed_worksheet.write(0, i + 1, killed_first_row[i][10::].replace("_", " ").title(), cell_first_column)
dropped_worksheet = workbook.add_worksheet("Dropped")
worksheets.append(dropped_worksheet)
for i in range(0, len(dropped_first_row)):
    dropped_worksheet.write(0, i + 1, dropped_first_row[i][10::].replace("_", " ").title(), cell_first_column)
picked_up_worksheet = workbook.add_worksheet("Picked up")
worksheets.append(picked_up_worksheet)
for i in range(0, len(picked_up_first_row)):
    picked_up_worksheet.write(0, i + 1, picked_up_first_row[i][10::].replace("_", " ").title(), cell_first_column)
custom_worksheet = workbook.add_worksheet("Custom")
worksheets.append(custom_worksheet)
for i in range(0, len(custom_first_row)):
    custom_worksheet.write(0, i + 1, custom_first_row[i][10::].replace("_", " ").title(), cell_first_column)
broken_worksheet = workbook.add_worksheet("Broken")
worksheets.append(broken_worksheet)
for i in range(0, len(broken_first_row)):
    broken_worksheet.write(0, i + 1, broken_first_row[i][10::].replace("_", " ").title(), cell_first_column)

# Filling up the spreadsheet
row = 1
for data in player_data:
    # Add name to each spreadsheet
    for sheet in worksheets:
        sheet.write(row, 0, player_uuid_name_dict.get(data[0]), cell_first_row)

    # Add statistics
    stats = data[1]["stats"]
    for i in range(0, len(killed_by_first_row)):
        killed_by_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:killed_by" in stats:
            if killed_by_first_row[i] in stats["minecraft:killed_by"]:
                killed_by_worksheet.write(row, i + 1, stats["minecraft:killed_by"][killed_by_first_row[i]], cell_stat)
    for i in range(0, len(crafted_first_row)):
        crafted_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:crafted" in stats:
            if crafted_first_row[i] in stats["minecraft:crafted"]:
                crafted_worksheet.write(row, i + 1, stats["minecraft:crafted"][crafted_first_row[i]], cell_stat)
    for i in range(0, len(mined_first_row)):
        mined_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:mined" in stats:
            if mined_first_row[i] in stats["minecraft:mined"]:
                mined_worksheet.write(row, i + 1, stats["minecraft:mined"][mined_first_row[i]], cell_stat)
    for i in range(0, len(used_first_row)):
        used_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:used" in stats:
            if used_first_row[i] in stats["minecraft:used"]:
                used_worksheet.write(row, i + 1, stats["minecraft:used"][used_first_row[i]], cell_stat)
    for i in range(0, len(killed_first_row)):
        killed_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:killed" in stats:
            if killed_first_row[i] in stats["minecraft:killed"]:
                killed_worksheet.write(row, i + 1, stats["minecraft:killed"][killed_first_row[i]], cell_stat)
    for i in range(0, len(dropped_first_row)):
        dropped_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:dropped" in stats:
            if dropped_first_row[i] in stats["minecraft:dropped"]:
                dropped_worksheet.write(row, i + 1, stats["minecraft:dropped"][dropped_first_row[i]], cell_stat)
    for i in range(0, len(picked_up_first_row)):
        picked_up_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:picked_up" in stats:
            if picked_up_first_row[i] in stats["minecraft:picked_up"]:
                picked_up_worksheet.write(row, i + 1, stats["minecraft:picked_up"][picked_up_first_row[i]], cell_stat)
    for i in range(0, len(custom_first_row)):
        custom_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:custom" in stats:
            if custom_first_row[i] in stats["minecraft:custom"]:
                if custom_first_row[i] == "minecraft:play_one_minute":
                    custom_worksheet.write(row, i + 1, stats["minecraft:custom"][custom_first_row[i]] / (20.0 * 60 * 60), cell_stat)
                    custom_worksheet.write(0, i + 1, "time played in hours".title(), cell_first_row)
                elif "time" in custom_first_row[i]:
                    custom_worksheet.write(row, i + 1, stats["minecraft:custom"][custom_first_row[i]] / (20.0 * 60), cell_stat)
                elif "one_cm" in custom_first_row[i]:
                    custom_worksheet.write(row, i + 1, stats["minecraft:custom"][custom_first_row[i]] / (100.0 * 1000), cell_stat)
                    custom_worksheet.write(0, i + 1, (custom_first_row[i][10:-2:] + "km").replace("_", " ").title(), cell_first_column)
                else:
                    custom_worksheet.write(row, i + 1, stats["minecraft:custom"][custom_first_row[i]], cell_stat)
    for i in range(0, len(broken_first_row)):
        broken_worksheet.write_blank(row, i + 1, "", cell_stat)
        if "minecraft:broken" in stats:
            if broken_first_row[i] in stats["minecraft:broken"]:
                broken_worksheet.write(row, i + 1, stats["minecraft:broken"][broken_first_row[i]], cell_stat)
    row += 1

# Freeze first column
for sheet in worksheets:
    sheet.freeze_panes(0, 1)

# Defining outer ranges
killed_by_outer_range = column_string(len(killed_by_first_row) + 1) + str(len(players) + 1)
crafted_outer_range = column_string(len(crafted_first_row) + 1) + str(len(players) + 1)
mined_outer_range = column_string(len(mined_first_row) + 1) + str(len(players) + 1)
used_outer_range = column_string(len(used_first_row) + 1) + str(len(players) + 1)
killed_outer_range = column_string(len(killed_first_row) + 1) + str(len(players) + 1)
dropped_outer_range = column_string(len(dropped_first_row) + 1) + str(len(players) + 1)
picked_up_outer_range = column_string(len(picked_up_first_row) + 1) + str(len(players) + 1)
custom_outer_range = column_string(len(custom_first_row) + 1) + str(len(players) + 1)
broken_outer_range = column_string(len(broken_first_row) + 1) + str(len(players) + 1)

# Custom conditional formatting: =B2=MAX(B$2:B$12) and alternating colors
killed_by_worksheet.conditional_format("B2:" + killed_by_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
killed_by_worksheet.conditional_format("B2:" + killed_by_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
crafted_worksheet.conditional_format("B2:" + crafted_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
crafted_worksheet.conditional_format("B2:" + crafted_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
mined_worksheet.conditional_format("B2:" + mined_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
mined_worksheet.conditional_format("B2:" + mined_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
used_worksheet.conditional_format("B2:" + used_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
used_worksheet.conditional_format("B2:" + used_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
killed_worksheet.conditional_format("B2:" + killed_by_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
killed_worksheet.conditional_format("B2:" + killed_by_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
dropped_worksheet.conditional_format("B2:" + dropped_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
dropped_worksheet.conditional_format("B2:" + dropped_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
picked_up_worksheet.conditional_format("B2:" + picked_up_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
picked_up_worksheet.conditional_format("B2:" + picked_up_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
custom_worksheet.conditional_format("B2:" + custom_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
custom_worksheet.conditional_format("B2:" + custom_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})
broken_worksheet.conditional_format("B2:" + broken_outer_range, {"type": "formula", "criteria": "=B2=MAX(B$2:B$" + str(len(players) + 1) + ")", "format": cell_light_magenta})
broken_worksheet.conditional_format("B2:" + broken_outer_range, {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": cell_light_pink})

# Adding autofilters
killed_by_worksheet.autofilter("B1:" + killed_by_outer_range)
crafted_worksheet.autofilter("B1:" + crafted_outer_range)
mined_worksheet.autofilter("B1:" + mined_outer_range)
used_worksheet.autofilter("B1:" + used_outer_range)
killed_worksheet.autofilter("B1:" + killed_outer_range)
dropped_worksheet.autofilter("B1:" + dropped_outer_range)
picked_up_worksheet.autofilter("B1:" + picked_up_outer_range)
custom_worksheet.autofilter("B1:" + custom_outer_range)
broken_worksheet.autofilter("B1:" + broken_outer_range)

workbook.close()
