################################################################
# Calculate and return list of decks and meta presence per deck.
def metagame(excel_file, t_date_1, t_date_2):
    ws = excel_file["Sheet1"]

    # Arbitrary upper limit for search purposes.
    TOURNAMENT_SIZE = 50

    # Prepare a list of tournaments.
    tournaments = []
    is_left_bound = False
    for cell in ws[1]:
        # Logic to check  area between two given dates.
        if t_date_1 in str(cell.value): is_left_bound = True
        if str(cell.value) != "None" and is_left_bound:
            tournaments.append(str(cell.value)[:10])
        if t_date_2 in str(cell.value): break

    decks = []
    i = 0   # Iterator for tournament reference.
    # Search the sheet for tournaments.
    for row in ws:
        for cell in row:
            if tournaments[i] in str(cell.value):
                j = 0
                while j < TOURNAMENT_SIZE:
                    # Add non-empty cells from that date to decks list.
                    if ws.cell(row=cell.row+2+j,
                               column=cell.column+2).value:
                        decks.append(ws.cell(row=cell.row+2+j,
                                             column=cell.column+2).value)
                    j += 1
                if i < len(tournaments)-1: i += 1

    # Convert decks into a Counter.
    from collections import Counter
    decks_count = Counter(decks)
    # Calculate and assign meta presence % to each deck.
    for deck in decks_count:
        decks_count[deck] = float(decks_count[deck] / len(decks) * 100)
    # Sort data in descending, alphabetical order.
    sort_dc_alpha = dict(sorted(decks_count.items()))
    meta_split = dict(sorted(sort_dc_alpha.items(), key=lambda item:item[1], reverse=True))

    return decks, meta_split

##################################################################################
# Separate winners and losers into lists, ignoring mirrors for decks. Draw = loss.
def deck_wr(excel_file, t_date_1, t_date_2):
    ws = excel_file["Sheet1"]

    # Prepare a list of tournaments.
    tournaments = []
    is_left_bound = False
    for cell in ws[1]:
        # Logic to check  area between two given dates.
        if t_date_1 in str(cell.value): is_left_bound = True
        if str(cell.value) != "None" and is_left_bound:
            tournaments.append(str(cell.value)[:10])
        if t_date_2 in str(cell.value): break

    # Prepare empty lists to populate.
    winners = []
    losers = []
    for date in tournaments:
        # Assign the correctly dated sheet.
        ws = excel_file[date]
        i = 1   # Iterator for rows reference.
        for row in ws:
            # Skip ROUND separator rows.
            if "ROUND" in str(ws['A'+str(i)].value):
                i += 1
                continue
            # Compare deck names, if mirror match, skip row.
            if ws['B'+str(i)].value == ws['E'+str(i)].value:
                i += 1
                continue
            # Compare score, if equal (draw), both are losers.
            if ws['C'+str(i)].value == ws['D'+str(i)].value:
                losers.append(ws['B'+str(i)].value)
                losers.append(ws['E'+str(i)].value)
                i += 1
                continue
            # Check C and D columns for winner and loser, append.
            if ws['C'+str(i)].value > ws['D'+str(i)].value:
                winners.append(ws['B'+str(i)].value)
                losers.append(ws['E'+str(i)].value)
            else:
                winners.append(ws['E'+str(i)].value)
                losers.append(ws['B'+str(i)].value)
            i += 1

    return winners, losers

######################################################
# Separate winners and losers into lists. Draw = loss.
def player_wr(excel_file, t_date_1, t_date_2):
    ws = excel_file["Sheet1"]

    # Prepare a list of tournaments.
    tournaments = []
    is_left_bound = False
    for cell in ws[1]:
        # Logic to check  area between two given dates.
        if t_date_1 in str(cell.value): is_left_bound = True
        if str(cell.value) != "None" and is_left_bound:
            tournaments.append(str(cell.value)[:10])
        if t_date_2 in str(cell.value): break

    # Prepare empty lists to populate.
    winners = []
    losers = []
    for date in tournaments:
        # Assign the correctly dated sheet.
        ws = excel_file[date]
        i = 1   # Iterator for rows reference.
        for row in ws:
            # Skip ROUND separator rows.
            if "ROUND" in str(ws['A'+str(i)].value):
                i += 1
                continue
            # Compare score, if equal (draw), both are losers.
            if ws['C'+str(i)].value == ws['D'+str(i)].value:
                losers.append(ws['A'+str(i)].value)
                losers.append(ws['F'+str(i)].value)
                i += 1
                continue
            # Check C and D columns for winner and loser, append.
            if ws['C'+str(i)].value > ws['D'+str(i)].value:
                winners.append(ws['A'+str(i)].value)
                losers.append(ws['F'+str(i)].value)
            else:
                winners.append(ws['F'+str(i)].value)
                losers.append(ws['A'+str(i)].value)
            i += 1

    return winners, losers

##################################################
# Calculate winrate from winners and losers lists.

def winrate(winners, losers):

    # Order lists into dictionaries.
    from collections import Counter
    winners_ord = Counter(winners)
    losers_ord = Counter(losers)

    # Prepare a dictionary (key:"PLAYER|DECK", value:WINRATE).
    winrate = {}
    # For every winner, find a corresponding loser.
    for winner in winners_ord:
        for loser in losers_ord:
            # If loser has no wins, winrate = 0.
            if loser not in winners: winrate[loser] = float(0)
            if winner != loser: continue
            # Calculate winrate (wins / games played).
            winrate[winner] = winners_ord[winner] * 100 / (
                            winners_ord[winner] + losers_ord[loser])
        # If winner has no losses, winrate = 100.
        if winner not in winrate:
            winrate[winner] = float(100)

    # Sort data in descending, alphabetical order.
    sort_wr_alpha = dict(sorted(winrate.items()))
    winrate = dict(sorted(sort_wr_alpha.items(), key=lambda item:item[1], reverse=True))

    return winrate