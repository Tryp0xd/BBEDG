import xlsxwriter
import requests
import math
import datetime
#Version control stuff
current_version = "0.0.2"

github_repo_url = "https://github.com/Tryp0xd/BBEDG"
try:
    
    response = requests.get(github_repo_url + "/releases/latest")
    latest_version = response.url.split("/")[-1]

    if current_version != latest_version:
        update_message = f"[UPDATE REQUIRED] BBEDG has been updated from {current_version} to {latest_version}, please go to {github_repo_url} to update this program. You may ignore this message."
        print(update_message)
    else:
        print("You are currently running the latest version of BBEDG!")
except (requests.exceptions.RequestException, requests.exceptions.ConnectionError):
    print("[ERROR] Failed to check versions, please ensure you have internet connection to see whether an update is available!")


def fetchBazaarInfo() -> dict:
    r = requests.get("https://api.hypixel.net/skyblock/bazaar")
    data = r.json()
    return data

def xlNotation(row, col):
    return xlsxwriter.utility.xl_rowcol_to_cell(row, col)





# Cell filling shit
start_row = 0
start_col = 0
end_row = 50
end_col = 50

relevantIDs = {"CHEAP_COFFEE"            : "Cheap Coffee",
               "DECENT_COFFEE"           : "Decent Coffee",
               "BLACK_COFFEE"            : "Black Coffee",
               "TEPID_GREEN_TEA"         : "Tepid Green Tea",
               "KNOCKOFF_COLA"           : "KnockOff™ Cola",
               "BITTER_ICE_TEA"          : "Bitter Ice Tea",
               "PULPOUS_ORANGE_JUICE"    : "Pulpous Orange Juice",
               "VIKING_TEAR"             : "Viking's Tear",
               "TUTTI_FRUTTI_POISON"     : "Tutti-Frutti Flavoured Poison",
               "DR_PAPER"                : "Dctr. Paper",
               "SLAYER_ENERGY_DRINK"     : "Slayer© Energy Drink",
               "TRUE_ESSENCE"            : "True Essence",
               "GOLDEN_CARROT"           : "Golden Carrot",
               "ENCHANTED_SUGAR_CANE"    : "Enchanted Sugar Cane",
               "ENCHANTED_COOKIE"        : "Enchanted Cookie",
               "ENCHANTED_RABBIT_FOOT"   : "Enchanted Rabbit Foot",
               "ENCHANTED_CAKE"          : "Enchanted Cake",
               "ENCHANTED_BLAZE_ROD"     : "Enchanted Blaze Rod",
               "ENCHANTED_GOLD_BLOCK"    : "Enchanted Gold Block",
               "FEATHER"                 : "Feather",
               "FLINT"                   : "Flint",
               "INK_SACK:4"              : "Lapis Lazuli",
               "ENCHANTED_COOKED_MUTTON" : "Enchanted Cooked Mutton",
               "ENCHANTED_CACTUS"        : "Enchanted Cactus",
               "COAL"                    : "Coal",
               "MITHRIL_ORE"             : "Mithril",
               "OBSIDIAN"                : "Obsidian",
               "RAW_FISH:1"              : "Raw Salmon",
               "ENCHANTED_RABBIT_HIDE"   : "Enchanted Rabbit Hide",
               "ENCHANTED_GHAST_TEAR"    : "Enchanted Ghast Tear",
               "ENCHANTED_PUFFERFISH"    : "Enchanted Pufferfish",
               "ENCHANTED_RED_SAND"      : "Enchanted Red Sand",
               "GREEN_CANDY"             : "Green Candy",
               "FULL_CHUM_BUCKET"        : "Full Chum Bucket",
               "GLOWING_MUSHROOM"        : "Glowing Mushroom",
               "SUGAR_CANE"              : "Sugar Cane",
               "BROWN_MUSHROOM"          : "Brown Mushroom",
               "SPIDER_EYE"              : "Spider Eye",
               "MAGMA_CREAM"             : "Magma Cream",
               "NETHER_STALK"            : "Nether Wart"}


# Set prices format : skyblockID : costInCoins
setPrices = {"CHEAP_COFFEE"         : 1000,
             "DECENT_COFFEE"        : 5000,
             "BLACK_COFFEE"         : 50000,
             "TEPID_GREEN_TEA"      : 1000,
             "KNOCKOFF_COLA"        : 1500,
             "BITTER_ICE_TEA"       : 1200,
             "PULPOUS_ORANGE_JUICE" : 1000,
             "VIKING_TEAR"          : 15000,
             "TUTTI_FRUTTI_POISON"  : 1000,
             "DR_PAPER"             : 4000,
             "SLAYER_ENERGY_DRINK"  : 10000,
             "TRUE_ESSENCE"         : 25000,
             "GOLDEN_CARROT"        : 15}


# Brewing product format : skyblockID : [potionName, rl, egb, erb, eg, nw]
# Lingo for last 5 components:
#                  rl  = Redstone Lamp
#                  egb = Enchanted Glowstone Block
#                  erb = Enchanted Redstone Block
#                  eg  = Enchanted Gunpowder
#                  nw  = Nether Wart
# No cost included as a different dictionary will be populated once
# we go to API part.

brewProd = {"ENCHANTED_SUGAR_CANE"    : ["Speed", 0, 1, 1, 1, 0],
            "ENCHANTED_COOKIE"        : ["Adrenaline", 0, 1, 1, 1, 0],
            "ENCHANTED_RABBIT_FOOT"   : ["Rabbit", 0, 1, 1, 1, 0],
            "ENCHANTED_CAKE"          : ["Agility", 0, 1, 1, 1, 0],
            "ENCHANTED_BLAZE_ROD"     : ["Strength", 0, 1, 1, 1, 0],
            "ENCHANTED_GOLD_BLOCK"    : ["Absorption", 0, 1, 1, 1, 0],
            "FEATHER"                 : ["Archery", 0, 1, 1, 1, 0],
            "FLINT"                   : ["Critical", 0, 1, 1, 1, 0],
            "INK_SACK:4"              : ["Experience", 0, 1, 1, 1, 0],
            "ENCHANTED_COOKED_MUTTON" : ["Mana", 0, 1, 1, 1, 0],
            "ENCHANTED_CACTUS"        : ["Resistance", 0, 1, 1, 1, 0],
            "COAL"                    : ["Haste", 0, 1, 1, 1, 1],
            "MITHRIL_ORE"             : ["Spelunker", 0, 1, 1, 1, 1],
            "OBSIDIAN"                : ["Stun", 0, 1, 1, 1, 1],
            "RAW_FISH:1"              : ["Dodge", 0, 1, 1, 1, 1],
            "TRUE_ESSENCE"            : ["True Resistance", 0, 1, 1, 1, 1],
            "ENCHANTED_RABBIT_HIDE"   : ["Pet Luck", 0, 1, 1, 1, 1],
            "ENCHANTED_GHAST_TEAR"    : ["Regeneration", 0, 1, 1, 1, 0],
            "ENCHANTED_PUFFERFISH"    : ["Water Breathing", 0, 1, 1, 1, 1],
            "ENCHANTED_RED_SAND"      : ["Burning", 0, 1, 1, 1, 1],
            "GOLDEN_CARROT"           : ["Night Vision", 0, 0, 1, 1, 1],
            "MAGMA_CREAM"             : ["Fire Resistance", 0, 0, 1, 1, 0]}

# This is to link what 'awkward potion' brew matches to what material

brewMaterialMatch = {
    "ENCHANTED_BLAZE_ROD"     : "KNOCKOFF_COLA",
    "ENCHANTED_GOLD_BLOCK"    : "DR_PAPER",
    "FEATHER"                 : "TUTTI_FRUTTI_POISON",
    "FLINT"                   : "SLAYER_ENERGY_DRINK",
    "INK_SACK:4"              : "VIKING_TEAR",
    "ENCHANTED_COOKED_MUTTON" : "BITTER_ICE_TEA",
    "ENCHANTED_CACTUS"        : "TEPID_GREEN_TEA",
    "ENCHANTED_GHAST_TEAR"    : "PULPOUS_ORANGE_JUICE"
}
brewList = [
"CHEAP_COFFEE",
"DECENT_COFFEE",
"BLACK_COFFEE",
"TEPID_GREEN_TEA",
"KNOCKOFF_COLA",
"BITTER_ICE_TEA",
"PULPOUS_ORANGE_JUICE",
"VIKING_TEAR",
"TUTTI_FRUTTI_POISON",
"DR_PAPER",
"SLAYER_ENERGY_DRINK"
]


# Potions which require a NPC purchase or require more than two brewing components
# Format : {*mats : amt, brewMat : [rl, egb, erb, eg, nw]}
otherBrewablePotions = {
    "Magic Find"         : {"materials"            : {"GREEN_CANDY": 10},
                            "brewMat"              : [0,1,1,1,0],
                            "resultType"           : 'potion'},
    "Spirit"             : {"materials"            : {"GREEN_CANDY": 1},
                            "brewMat"              : [0,1,1,1,0],
                            "resultType"           : 'potion'},
    "Mushed Glowy Tonic" : {"materials"            : {"FULL_CHUM_BUCKET" : 8,
                                                      "GLOWING_MUSHROOM" : 10},
                            "brewMat"              : [0,0,0,1,0],
                            "resultType"           : 'potion'},
    "Invisibility"       : {"materials"            : {"GOLDEN_CARROT"  : 1,
                                                      "SUGAR_CANE"     : 1,
                                                      "BROWN_MUSHROOM" : 1,
                                                      "SPIDER_EYE"     : 1},
                            "brewMat"              : [0,1,1,1,1],
                            "resultType"           : 'item'}
}
otherBrewableMats = {}

for n, vals in enumerate(otherBrewablePotions.items()):
    
    name, data = vals
    mats = data['materials']
    for z, vals2 in enumerate(mats.items()):
        id, count = vals2
        id = str(id)  # Convert id to string
        if id not in otherBrewableMats:
            if data['resultType'] == 'potion':
                otherBrewableMats[id] = count*3
            else:
                otherBrewableMats[id] = count
        else:
            getCount = otherBrewableMats[id]
            if data['resultType'] == 'potion':
                otherBrewableMats[id] = getCount + count*3
            else:
                otherBrewableMats[id] = getCount + count

# assuming only rl 
xpPotions = ["Alchemy", "Combat", "Enchanting", "Farming", "Fishing", "Foraging", "Mining"]

# Miscalleanous potions (from AH, ask user each potion in here for their average price)
# Format : {potion : [rl, egb, erb, eg, nw]}
miscPotions = {"Wisp" : [0, 0, 1, 0, 0]}
miscPotPrice = {}

# Create a excel file and adds a worksheet.
wb = xlsxwriter.Workbook('Splasher Spreadsheet.xlsx')
ws = wb.add_worksheet('Brewing Guide')
ws2 = wb.add_worksheet('Cost Summary')
fmt = wb.add_format()

tb_options = {'width' : 363,
              'height' : 100,
              'x_offset' : 0,
              'y_offset' : 0,
              }

dm_textbox = {'color' : 'black'}
lm_textbox = {'color' : 'white'}

# Define the header labels
header_labels = ['Potion Name', 'Item Name(s)', 'RL', 'EGB', 'ERB', 'EG', 'NW', 'Brew']
header_labels_summary = ['Set amount (3x)', '', 'Item', 'Qty', 'Buy Order', 'Instant Buy', 'Buy Order x Qty', 'Instant Buy x Qty']

# Set widths
ws.set_column("A:A", 18)
ws.set_column("B:B", 65)
ws.set_column("C:C", 2)
ws.set_column("D:D", 3)
ws.set_column("E:E", 3)
ws.set_column("F:F", 2)
ws.set_column("G:G", 3)
ws.set_column("H:H", 23)



# Set style mode
mode = input("What mode do you want the spreadsheet to be?\nLight mode - LM\nDark mode - DM\nOption : ").lower()
if mode == "lm":
    fmt.set_bg_color("#FFFFFF")
    tb_options['font'] = {'color' : 'black',
                          'size'  : 14}
    tb_options['fill'] = {'color' : 'black'}
elif mode == "dm":
    fmt.set_bg_color("#000000")
    fmt.set_font_color("white")
    tb_options['font'] = {'color' : 'white',
                          'size'  : 14}
    tb_options['fill'] = {'color' : 'black'}
else:
    print("You chose none of the options, defaulting to regular excel mode (sorry not sorry for singing your eyeballs)")
    tb_options['font'] = {'color' : 'black',
                          'size'  : 14}
    tb_options['fill'] = {'color' : 'white'}

fmt.set_num_format('#,##0')
coffeeTypeChoice = int(input("What coffee type?\n1 - Cheap\n2 - Decent\n3 - Black\nOption : "))
if coffeeTypeChoice < 1 and coffeeTypeChoice > 3:
    print("Defaulting coffee to cheap because you didnt bother picking the right number")
    coffeeType = 1
coffeeType = ""
match coffeeTypeChoice:
    case 1:
        coffeeType = "CHEAP_COFFEE"
    case 2:
        coffeeType = "DECENT_COFFEE"
    case 3:
        coffeeType = "BLACK_COFFEE"

for n, vals in enumerate(miscPotions.items()):
    name, brewMat = vals
    cost = input(f"{name} seems to be a miscellanous item! How much do you usually pay for this? (whole number) : ")
    if cost.isnumeric() != True:
        cost = 0
        print("i said whole numbers bozo, set cost to 0.")
    miscPotPrice[name] = int(cost)

# [ ---------------------------------------------[Brewing Guide]--------------------------------------------- ]

# Fill the worksheet based on user preference.
for row in range(start_row, end_row + 1):
    for col in range(start_col, end_col + 1):
        ws.write(row, col, '', fmt)

# Write the header labels to the worksheet
for col, label in enumerate(header_labels):
    ws.write(0, col, label, fmt)

# Fill material pot data, then display data
format_good = wb.add_format({'bg_color': 'green', 'font_color' : 'green'})
format_bad = wb.add_format({'bg_color': 'red', 'font_color': 'red'})

curr_row = 0
for row, values in enumerate(brewProd.items()):
    id, setting = values
    individualItem = setting
    individualItem.insert(1, relevantIDs[id])
    for n in range(len(individualItem)):
        ws.write(row+1, n, setting[n], fmt)
        if n >= 2 and n <= 6:
            if setting[n] == 0:
                ws.conditional_format(row+1, n, row+1, n, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': format_bad})
            else:
                ws.conditional_format(row+1, n, row+1, n, {'type': 'cell', 'criteria': '=', 'value': 1, 'format': format_good})
        if n==6:
            if setting[n] == 0:
                coffeePots = ["ENCHANTED_SUGAR_CANE", "ENCHANTED_COOKIE", "ENCHANTED_RABBIT_FOOT", "ENCHANTED_CAKE"]
                if id in coffeePots:
                    ws.write(row+1, n+1, relevantIDs[coffeeType], fmt)
                elif id in brewMaterialMatch:
                    ws.write(row+1, n+1, relevantIDs[brewMaterialMatch[id]], fmt)
                else:
                    ws.write(row+1, n+1, "N/A", fmt)
            else:
                ws.write(row+1, n+1, "Awkward Potion", fmt)
    curr_row = row+2
       
for n, values in enumerate(otherBrewablePotions.items()):
    potionName, setting = values
    individualItem = [potionName]
    brewMats = []
    for name, item in setting.items():
        if name == "materials":
            materials = item
            material_string = ", ".join([f"{quantity}x {relevantIDs[material]}" for material, quantity in materials.items()])
        elif name == "brewMat":
            brewMats = item
        else:
            continue

    individualItem.append(material_string)
    for item in brewMats:
        individualItem.append(item)

    

    for n in range(len(individualItem)):      
        ws.write(curr_row, n, individualItem[n], fmt)
        if n >= 2 and n <= 6:
            if individualItem[n] == 0:
                ws.conditional_format(curr_row, n, curr_row, n, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': format_bad})
            else:
                ws.conditional_format(curr_row, n, curr_row, n, {'type': 'cell', 'criteria': '=', 'value': 1, 'format': format_good})
        if n==6:
            if individualItem[n] == 0:
                    ws.write(curr_row, n+1, "N/A", fmt)
            else:
                ws.write(curr_row, n+1, "Awkward Potion", fmt)
        
    curr_row += 1

for n, name in enumerate(xpPotions):
    xpPotArr = [name, name+" XP Boost", 1, 0, 0, 1, 0]
    for k in range(len(xpPotArr)):
        ws.write(curr_row, k, xpPotArr[k], fmt)
        if k >= 2 and k <= 6:
            if xpPotArr[k] == 0:
                ws.conditional_format(curr_row, k, curr_row, k, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': format_bad})
            else:
                ws.conditional_format(curr_row, k, curr_row, k, {'type': 'cell', 'criteria': '=', 'value': 1, 'format': format_good})
        if k==6:
            ws.write(curr_row, k+1, "N/A", fmt)
    curr_row += 1
    

for n, vals in enumerate(miscPotions.items()):
    name, brewMat = vals
    miscArr = [name, name]
    for item in brewMat:
        miscArr.append(item)
    for k in range(len(miscArr)):
        ws.write(curr_row, k, miscArr[k], fmt)
        if k >= 2 and k <= 6:
            if miscArr[k] == 0:
                ws.conditional_format(curr_row, k, curr_row, k, {'type': 'cell', 'criteria': '=', 'value': 0, 'format': format_bad})
            else:
                ws.conditional_format(curr_row, k, curr_row, k, {'type': 'cell', 'criteria': '=', 'value': 1, 'format': format_good})
        if k==6:
            ws.write(curr_row, k+1, "N/A", fmt)
    curr_row += 1

# [ ---------------------------------------------[Item Cost Summary]--------------------------------------------- ]

for row in range(start_row, end_row + 1):
    for col in range(start_col, end_col + 1):
        ws2.write(row, col, '', fmt)

for col, label in enumerate(header_labels_summary):
    ws2.write(0, col, label, fmt)

bz_content = fetchBazaarInfo()

ws2.set_column("A:A", 15)
ws2.set_column("C:C", 23)
ws2.set_column("D:D", 3)
ws2.set_column("E:E", 12)
ws2.set_column("F:F", 12)
ws2.set_column("G:G", 15)
ws2.set_column("H:H", 15)

ws2.write(1,0, 1, fmt) # Writes '1' in Set amount.

item_col = 2
item_row = 1

qty = 1
b_order = 2
i_buy = 3
b_order_x_qty = 4
i_buy_x_qty = 5

nw_count = 0
itemQty = 0

for n, valx in enumerate(otherBrewablePotions.items()):
    name, setting = valx
    if setting['brewMat'][-1] == 1:
        nw_count += 1


for row, vals in enumerate(relevantIDs.items()):
    id, name = vals
    ws2.write(item_row+row, item_col, name, fmt)
    
    if id in brewList:
        itemQty = 3
        if name in ["Cheap Coffee", "Black Coffee", "Decent Coffee"]:
            if id != coffeeType:
                itemQty = 0
            else:
                itemQty *= 4
    elif id in brewProd:
        if brewProd[id][-1] == 1:
            nw_count += 1
        itemQty = 1
    elif id in otherBrewableMats:
        itemQty = otherBrewableMats[id]
    elif id == "NETHER_STALK":
        itemQty = nw_count
    else:
        itemQty = 1


    ws2.write(item_row+row, item_col+qty, f"=A2*{itemQty}", fmt)
    


    if id in setPrices:
        if name in ["Cheap Coffee", "Decent Coffee", "Black Coffee"]:
            if id != coffeeType:
                ws2.write(item_row+row, item_col+b_order, 0, fmt)
                ws2.write(item_row+row, item_col+i_buy, 0, fmt)
                ws2.write(item_row+row, item_col+b_order_x_qty, 0, fmt)
                ws2.write(item_row+row, item_col+i_buy_x_qty, 0, fmt)
            else:
                ws2.write(item_row+row, item_col+b_order, setPrices[id], fmt)
                ws2.write(item_row+row, item_col+i_buy, setPrices[id], fmt)   
                ws2.write(item_row+row, item_col+b_order_x_qty, f"={xlNotation(item_row+row, item_col+qty)}*{xlNotation(item_row+row, item_col+b_order)}", fmt)
                ws2.write(item_row+row, item_col+i_buy_x_qty, f"={xlNotation(item_row+row, item_col+qty)}*{xlNotation(item_row+row, item_col+i_buy)}", fmt)   
        else:
                ws2.write(item_row+row, item_col+b_order, setPrices[id], fmt)
                ws2.write(item_row+row, item_col+i_buy, setPrices[id], fmt)
                ws2.write(item_row+row, item_col+b_order_x_qty, f"={xlNotation(item_row+row, item_col+qty)}*{setPrices[id]}", fmt)
                ws2.write(item_row+row, item_col+i_buy_x_qty, f"={xlNotation(item_row+row, item_col+qty)}*{setPrices[id]}", fmt)
    elif id in bz_content["products"]:
        buy_order = bz_content["products"][id]["quick_status"]["sellPrice"]
        instant_buy = bz_content["products"][id]["quick_status"]["buyPrice"]
        ws2.write(item_row+row, item_col+b_order, round(buy_order, 1), fmt)
        ws2.write(item_row+row, item_col+i_buy, round(instant_buy, 1), fmt)
        ws2.write(item_row+row, item_col+b_order_x_qty, f"=ROUND({xlNotation(item_row+row, item_col+qty)}*{xlNotation(item_row+row, item_col+b_order)}, 1)", fmt)
        ws2.write(item_row+row, item_col+i_buy_x_qty, f"=ROUND({xlNotation(item_row+row, item_col+qty)}*{xlNotation(item_row+row, item_col+i_buy)}, 1)", fmt)
    else:
        continue
item_row += len(relevantIDs)

for n, vals in enumerate(miscPotPrice.items()):
    name, cost = vals
    ws2.write(item_row+n, item_col, name, fmt)
    ws2.write(item_row+n, item_col+qty, "=A2*3", fmt)
    ws2.write(item_row+n, item_col+b_order, cost, fmt)
    ws2.write(item_row+n, item_col+i_buy, cost, fmt)
    ws2.write(item_row+n, item_col+b_order_x_qty, f"={xlNotation(item_row+n, item_col+qty)}*{xlNotation(item_row+n, item_col+b_order)}", fmt)
    ws2.write(item_row+n, item_col+i_buy_x_qty, f"={xlNotation(item_row+n, item_col+qty)}*{xlNotation(item_row+n, item_col+i_buy)}", fmt)

get_time_unix = int(bz_content['lastUpdated'])/ 1000
get_date = datetime.datetime.fromtimestamp(get_time_unix)
formatted_date = get_date.strftime('%Y-%m-%d %H:%M:%S')

tb_text = f'You may not see every item listed, so you need to scroll down.\nBazaar prices are as of {formatted_date}\nIf time doesnt match, its likely UK time idk'


# [ ---------------------------------------------[Material Cost Summary]--------------------------------------------- ]

ws2.insert_textbox('K1', tb_text, tb_options)

summary_qty = 1
summary_bo = 2
summary_ib = 3
summary_boq = 4
summary_ibq = 5

summary_headers = ['Material', 'Qty', 'Buy Order', 'Instant Buy', 'Buy Order x Qty', 'Instant Buy x Qty']
summary_row_ids = {"ENCHANTED_REDSTONE_LAMP"   : "Enchanted Redstone Lamp",
                   "ENCHANTED_GLOWSTONE"       : "Enchanted Glowstone Block",
                   "ENCHANTED_REDSTONE_BLOCK"  : "Enchhanted Redstone Block", 
                   'ENCHANTED_GUNPOWDER'       : "Enchanted Gunpowder"}



ws2.set_column("K:K", 25)
ws2.set_column("L:L",  3)
ws2.set_column(4, 7, 15)
ws2.set_column(12, 15, 15)

# K6 - headers
header_col = 10
header_row = 6
for col, label in enumerate(summary_headers):
    ws2.write(header_row-1, header_col+col, label, fmt)


temp = 0
for row, items in enumerate(summary_row_ids.items()):
    id, name = items
    buy_order = bz_content["products"][id]["quick_status"]["sellPrice"]
    instant_buy = bz_content["products"][id]["quick_status"]["buyPrice"]
    ws2.write(header_row+row, header_col, name, fmt)
    ws2.write(header_row+row, header_col+summary_qty, f"=SUM('Brewing Guide'!{xlNotation(1, 2+row)}:{xlNotation(35, 2+row)})*A2",fmt)
    ws2.write(header_row+row, header_col+summary_bo, round(buy_order,1), fmt)
    ws2.write(header_row+row, header_col+summary_ib, round(instant_buy,1), fmt)
    ws2.write(header_row+row, header_col+summary_boq, f"={xlNotation(header_row+row, header_col+summary_qty)}*{xlNotation(header_row+row, header_col+summary_bo)}", fmt)
    ws2.write(header_row+row, header_col+summary_ibq, f"={xlNotation(header_row+row, header_col+summary_qty)}*{xlNotation(header_row+row, header_col+summary_ib)}", fmt)
    temp = row

header_row += temp
header_row+=2

ws2.write(header_row, header_col, "Material Cost", fmt)
ws2.write(header_row, header_col+summary_bo, "=SUM(E2:E42)",fmt)
ws2.write(header_row, header_col+summary_ib, "=SUM(F2:F42)",fmt)
ws2.write(header_row, header_col+summary_boq, "=SUM(G2:G42)",fmt)
ws2.write(header_row, header_col+summary_ibq, "=SUM(H2:H42)",fmt)
ws2.write(header_row+1, header_col, "Sum", fmt)
ws2.write(header_row+1, header_col+summary_bo, "=SUM(M7:M12)",fmt)
ws2.write(header_row+1, header_col+summary_ib, "=SUM(N2:N12)",fmt)
ws2.write(header_row+1, header_col+summary_boq, "=SUM(O2:O12)",fmt)
ws2.write(header_row+1, header_col+summary_ibq, "=SUM(P2:P12)",fmt)

monetary_format = wb.add_format({'num_format': '#,##0.00'})

wb.close()
