########################################
# Convert date range into list of dates.
def tournament_dates(excel_file, t_date_1, t_date_2):
    ws = excel_file["Sheet1"]

    # Prepare an empty list and a boolean for bounding.
    date_list = []
    is_left_bound = False
    for cell in ws[1]:
        # Logic to check area between two given dates.
        if t_date_1 in str(cell.value): is_left_bound = True
        if str(cell.value) != "None" and is_left_bound:
            date_list.append(str(cell.value)[:10])
        if t_date_2 in str(cell.value): break
    
    return date_list

################################################################
# Calculate and return list of decks and meta presence per deck.
def metagame(excel_file, date_list):
    ws = excel_file["Sheet1"]

    # Arbitrary upper limit for search purposes.
    TOURNAMENT_SIZE = 100

    decks = []
    i = 0   # Iterator for tournament reference.
    # Search the sheet for tournaments.
    for row in ws:
        for cell in row:
            if date_list[i] in str(cell.value):
                j = 0
                while j < TOURNAMENT_SIZE:
                    # Add non-empty cells from that date to decks list.
                    deck = ws.cell(row=cell.row+2+j, column=cell.column+2).value
                    if deck:
                        # Remove the brackets for minor deck differences.
                        if deck.find('(') != -1:
                            decks.append(deck[:deck.find('(')-1])
                        else: decks.append(deck)
                    j += 1
                if i < len(date_list)-1: i += 1

    # Convert decks into a Counter.
    from collections import Counter
    decks_count = Counter(decks)
    # Calculate and assign meta presence % to each deck.
    unsort_dc = {}
    for deck in decks_count:
        unsort_dc[deck] = float(decks_count[deck] / len(decks) * 100)
    # Sort data in descending, alphabetical order.
    sort_dc_alpha = dict(sorted(unsort_dc.items()))
    meta_percentage = dict(sorted(sort_dc_alpha.items(), key=lambda item:item[1], reverse=True))
    # Populate a list with meta presence ratios.
    meta_ratio = []
    for deck in meta_percentage:
        meta_ratio.append([deck, decks_count[deck], len(decks)])

    return decks, meta_percentage, meta_ratio

##################################################################################
# Separate winners and losers into lists, ignoring mirrors for decks. Draw = loss.
def deck_wr(excel_file, date_list):
    ws = excel_file["Sheet1"]

    # Prepare empty lists to populate.
    winners = []
    losers = []
    for date in date_list:
        # Assign the correctly dated sheet.
        ws = excel_file[date]
        i = 1   # Iterator for rows reference.
        for row in ws:    
            deck_left = str(ws['B'+str(i)].value)
            deck_right = str(ws['E'+str(i)].value)
            # Remove the brackets for minor deck differences.
            if deck_left.find('(') != -1:
                deck_left = deck_left[:deck_left.find('(')-1]
            if deck_right.find('(') != -1:
                deck_right = deck_right[:deck_right.find('(')-1]
            # Skip ROUND separator rows.
            if "ROUND" in str(ws['A'+str(i)].value):
                i += 1
                continue
            # Compare deck names, if mirror match, skip row.
            if deck_left == deck_right:
                i += 1
                continue
            # Compare score, if equal (draw), both are losers.
            if ws['C'+str(i)].value == ws['D'+str(i)].value:
                losers.append(deck_left)
                losers.append(deck_right)
                i += 1
                continue
            # Check C and D columns for winner and loser, append.
            if ws['C'+str(i)].value > ws['D'+str(i)].value:
                winners.append(deck_left)
                losers.append(deck_right)
            else:
                winners.append(deck_right)
                losers.append(deck_left)
            i += 1

    return winners, losers

######################################################
# Separate winners and losers into lists. Draw = loss.
def player_wr(excel_file, date_list):
    ws = excel_file["Sheet1"]

    # Prepare empty lists to populate.
    winners = []
    losers = []
    for date in date_list:
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

    winrate_percentage = {} # Dictionary (key:"PLAYER|DECK", value:WINRATE)
    winrate_ratio = [] # List ([PLAYER|DECK, WINS, MATCHES])
    # For every winner, find a corresponding loser.
    for winner in winners_ord:
        for loser in losers_ord:
            # If loser has no wins, winrate = 0.
            if loser not in winners:
                winrate_percentage[loser] = float(0)
                winrate_ratio.append([loser, 0, losers_ord[loser]])
            if winner != loser: continue
            # Calculate winrate (wins / games played).
            winrate_percentage[winner] = winners_ord[winner] * 100 / (
                                         winners_ord[winner] + losers_ord[loser])
            winrate_ratio.append([winner,
                                  winners_ord[winner],
                                  winners_ord[winner] + losers_ord[loser]])
        # If winner has no losses, winrate = 100.
        if winner not in winrate_percentage:
            winrate_percentage[winner] = float(100)
            winrate_ratio.append([winner,
                                  winners_ord[winner],
                                  winners_ord[winner]])

    # Sort dict data in descending, alphabetical order.
    sort_wr_alpha = dict(sorted(winrate_percentage.items()))
    winrate_percentage = dict(sorted(sort_wr_alpha.items(), key=lambda item:item[1], reverse=True))

    return winrate_percentage, winrate_ratio

####################################################
# Parse and establish a matchup matrix given a deck.

def matchup_matrix(excel_file, date_list, deck_name):
    ws = excel_file["Sheet1"]
   
    # Prepare empty dictionary to populate.
    matchup_matrix = {}
    for date in date_list:
        # Assign the correctly dated sheet.
        ws = excel_file[date]
        i = 1   # Iterator for rows reference.
        for row in ws:
            deck_left = str(ws['B'+str(i)].value)
            deck_right = str(ws['E'+str(i)].value)
            # Remove the brackets for minor deck differences.
            if deck_left.find('(') != -1:
                deck_left = deck_left[:deck_left.find('(')-1]
            if deck_right.find('(') != -1:
                deck_right = deck_right[:deck_right.find('(')-1]
            # Check only if matchup IS NOT mirror.
            if deck_left != deck_right:
                if deck_left == deck_name:
                    # Count deck left-hand side wins.
                    if ws['C'+str(i)].value > ws['D'+str(i)].value:
                        try: matchup_matrix[deck_right][0] += 1
                        except: matchup_matrix[deck_right] = [1, 0]
                    # Count deck left-hand side losses.
                    elif ws['C'+str(i)].value < ws['D'+str(i)].value:
                        try: matchup_matrix[deck_right][1] += 1
                        except: matchup_matrix[deck_right] = [0, 1]
                if deck_right == deck_name:
                    # Count deck right-hand side wins.
                    if ws['C'+str(i)].value < ws['D'+str(i)].value:
                        try: matchup_matrix[deck_left][0] += 1
                        except: matchup_matrix[deck_left] = [1, 0]
                    # Count deck right-hand side losses.
                    elif ws['C'+str(i)].value > ws['D'+str(i)].value:
                        try: matchup_matrix[deck_left][1] += 1
                        except: matchup_matrix[deck_left] = [0, 1]                    
            i += 1
    
    #! Sort data before returning.
    #! Alphabetically? Custom?

    return matchup_matrix