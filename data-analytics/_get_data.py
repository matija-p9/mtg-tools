import calculate as calc
import openpyxl

excel_file = openpyxl.load_workbook("./reference/liga2.xlsx")
date_range = ["2024-05-08", "2024-05-08"]

# Get metagame split per deck.
decks, meta_split = calc.metagame(excel_file, date_range[0], date_range[1])
# Print metagame split per deck.
print()
print("METAGAME SPLIT")
print("--------------")
for key in meta_split:
    print(f"{key}: {meta_split[key]:.2f}")

# Get winrate per player.
winners, losers = calc.player_wr(excel_file, date_range[0], date_range[1])
winrate = calc.winrate(winners, losers)
# Print winrate per player.
print()
print("PLAYER WINRATES")
print("---------------")
for key in winrate:
    print(f"{key}: {winrate[key]:.2f}")

# Get winrate per deck.
winners, losers = calc.deck_wr(excel_file, date_range[0], date_range[1])
winrate = calc.winrate(winners, losers)
# Print winrate per deck.
print()
print("DECK WINRATES")
print("-------------")
for key in winrate:
    print(f"{key}: {winrate[key]:.2f}")