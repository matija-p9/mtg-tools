tournament = "2024-04-24"
import calculate as calc
import openpyxl
wb = openpyxl.load_workbook("./reference/liga2.xlsx")
ws = wb[tournament]

# Get winrate per player.
winners, losers = calc.player_wr(ws)
winrate = calc.winrate(winners, losers)
# Print winrate per player.
print()
print("PLAYER WINRATES")
print("---------------")
for key in winrate:
    print(f"{key}: {winrate[key]:.2f}")

#######################################

# Get winrate per deck.
winners, losers = calc.deck_wr(ws)
winrate = calc.winrate(winners, losers)
# Print winrate per deck.
print()
print("DECK WINRATES")
print("-------------")
for key in winrate:
    print(f"{key}: {winrate[key]:.2f}")

#######################################

# Get metagame split per deck.
decks, meta_split = calc.metagame(wb["Sheet1"], tournament)
# Print metagame split per deck.
print()
print("METAGAME SPLIT")
print("--------------")
for key in meta_split:
    print(f"{key}: {meta_split[key]:.2f}")