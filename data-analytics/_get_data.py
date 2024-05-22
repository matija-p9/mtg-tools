import calculate as calc
import openpyxl

EXCEL_FILE = openpyxl.load_workbook("./reference/MTG Pula - Liga 2024 #1.xlsx")
DATE_RANGE = ["2024-05-15", "2024-04-24"]

date_list = calc.tournament_dates(EXCEL_FILE, DATE_RANGE[0], DATE_RANGE[1])

'''''''''''''''''''''
GET METAGAME SPLIT PER DECK.
'''''''''''''''''''''

# Function call to process data.
decks, meta_percentage, meta_ratio = calc.metagame(EXCEL_FILE, date_list)
winrate_ratio = {}
i = 0
# Convert data to dict for use.
for deck in meta_ratio:
    winrate_ratio[meta_ratio[i][0]] = meta_ratio[i][1]
    i += 1
# Print metagame split per deck.
print()
print("METAGAME SPLIT")
print("--------------")
for deck in meta_percentage:
    print(f"{deck} | {meta_percentage[deck]:.2f}% | {winrate_ratio[deck]}/{meta_ratio[0][2]}")

'''''''''''''''''''''
GET WINRATE PER PLAYER.
'''''''''''''''''''''

# Function call to process data.
winners, losers = calc.player_wr(EXCEL_FILE, date_list)
winrate_percentage, wr_ratio_list = calc.winrate(winners, losers)
winrate_ratio = {}
i = 0
# Convert data to dict for use.
for player in wr_ratio_list:
    winrate_ratio[wr_ratio_list[i][0]] = {"W": str(wr_ratio_list[i][1]), "M": str(wr_ratio_list[i][2])}
    i += 1
# Print winrate per player.
print()
print("PLAYER WINRATES")
print("---------------")
for player in winrate_percentage:
    print(f"{player} | {winrate_percentage[player]:.2f}% | {winrate_ratio[player]["W"]}/{winrate_ratio[player]["M"]}")

'''''''''''''''''''''
GET WINRATE PER DECK.
'''''''''''''''''''''

# Function calls to process data.
winners, losers = calc.deck_wr(EXCEL_FILE, date_list)
winrate_percentage, wr_ratio_list = calc.winrate(winners, losers)
winrate_ratio = {}
i = 0
# Convert data to dict for use.
for deck in wr_ratio_list:
    winrate_ratio[wr_ratio_list[i][0]] = {"W": str(wr_ratio_list[i][1]), "M": str(wr_ratio_list[i][2])}
    i += 1
# Print winrate per deck.
print()
print("DECK WINRATES")
print("-------------")
for deck in winrate_percentage:
    print(f"{deck} | {winrate_percentage[deck]:.2f}% | {winrate_ratio[deck]["W"]}/{winrate_ratio[deck]["M"]}")

'''''''''''''''''''''
GET MATCHUP MATRIX.
'''''''''''''''''''''

# Iterate through decks.
matrix_raw = []
for deck in winrate_percentage:
    # Function call to get matchup data.
    matrix_deck = calc.matchup_matrix(EXCEL_FILE, date_list, deck)
    for matchup in matrix_deck:
        matrix_raw.append([deck, matchup, matrix_deck[matchup]])
# Sort data alphabetically.
from operator import itemgetter
matrix = sorted(matrix_raw, key=itemgetter(0,1))
# Print whole matchup matrix.
deck_previous = ""
i = 0
for item in matrix:
    deck = matrix[i][0]
    if deck != deck_previous:
        print(f"\n{deck.upper()} MATRIX")
        print("-" * len(f"{deck.upper()} MATRIX"))
        deck_previous = deck
    print(f"{deck} {matrix[i][2][0]}-{matrix[i][2][1]} {matrix[i][1]} | {matrix[i][2][0]/(matrix[i][2][0]+matrix[i][2][1])*100:.2f}%")
    i += 1