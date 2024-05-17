import calculate as calc
import openpyxl
excel_file = openpyxl.load_workbook("./reference/MTG Pula - Liga 2024 #1.xlsx")
date_range = ["2024-05-15", "2024-04-24"]

'''
Get metagame split per deck.
'''
decks, meta_percentage, meta_ratio = calc.metagame(excel_file, date_range[0], date_range[1])
ratio_dict = {}
i = 0
# Convert data to dict for use.
for deck in meta_ratio:
    ratio_dict[meta_ratio[i][0]] = meta_ratio[i][1]
    i += 1
# Print metagame split per deck.
print()
print("METAGAME SPLIT")
print("--------------")
for deck in meta_percentage:
    print(f"{deck} | {meta_percentage[deck]:.2f}% | {ratio_dict[deck]}/{meta_ratio[0][2]}")

'''
Get winrate per player.
'''
winners, losers = calc.player_wr(excel_file, date_range[0], date_range[1])
winrate_percentage, winrate_ratio = calc.winrate(winners, losers)
ratio_dict = {}
i = 0
# Convert data to dict for use.
for player in winrate_ratio:
    ratio_dict[winrate_ratio[i][0]] = str(winrate_ratio[i][1]) + "/" + str(winrate_ratio[i][2])
    i += 1
# Print winrate per player.
print()
print("PLAYER WINRATES")
print("---------------")
for player in winrate_percentage:
    print(f"{player} | {winrate_percentage[player]:.2f}% | {ratio_dict[player]}")

'''
Get winrate per deck.
'''
winners, losers = calc.deck_wr(excel_file, date_range[0], date_range[1])
winrate_percentage, winrate_ratio = calc.winrate(winners, losers)
ratio_dict = {}
i = 0
# Convert data to dict for use.
for deck in winrate_ratio:
    ratio_dict[winrate_ratio[i][0]] = str(winrate_ratio[i][1]) + "/" + str(winrate_ratio[i][2])
    i += 1
# Print winrate per deck.
print()
print("DECK WINRATES")
print("-------------")
for deck in winrate_percentage:
    print(f"{deck} | {winrate_percentage[deck]:.2f}% | {ratio_dict[deck]}")