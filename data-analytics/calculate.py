################################################################
# Calculate and return list of decks and meta presence per deck.
def metagame(excel_worksheet, tournament_date):

    TOURNAMENT_SIZE = 50    # Arbitrary upper limit for purposes of searching.
    decks = []
    # Search the sheet for tournament_date.
    for row in excel_worksheet:
        for cell in row:
            if tournament_date in str(cell.value):
                i = 0
                while i < TOURNAMENT_SIZE:
                    # Add non-empty cells from that date to decks list.
                    if excel_worksheet.cell(row=cell.row+2+i,
                                            column=cell.column+2).value:
                        decks.append(excel_worksheet.cell(row=cell.row+2+i,
                                                          column=cell.column+2).value)
                    i += 1

    # Convert decks into a Counter.
    from collections import Counter
    decks_count = Counter(decks)
    # Calculate and assign meta presence % to each deck.
    for deck in decks_count:
        decks_count[deck] = float(decks_count[deck] / len(decks) * 100)
    # Sort data in descending order.
    meta_split = dict(sorted(decks_count.items(), key=lambda item:item[1], reverse=True))

    return decks, meta_split

##################################################################################
# Separate winners and losers into lists, ignoring mirrors for decks. Draw = loss.
def deck_wr(excel_worksheet):

    winners = []
    losers = []
    i = 1   # Iterator for rows reference.
    for row in excel_worksheet:
        # Skip ROUND separator rows.
        if "ROUND" in str(excel_worksheet['A'+str(i)].value):
            i += 1
            continue
        # Compare deck names, if mirror match, skip row.
        if excel_worksheet['B'+str(i)].value == excel_worksheet['E'+str(i)].value:
            i += 1
            continue
        # Compare score, if equal (draw), both are losers.
        if excel_worksheet['C'+str(i)].value == excel_worksheet['D'+str(i)].value:
            losers.append(excel_worksheet['B'+str(i)].value)
            losers.append(excel_worksheet['E'+str(i)].value)
            i += 1
            continue
        # Check C and D columns for winner and loser, append.
        if excel_worksheet['C'+str(i)].value > excel_worksheet['D'+str(i)].value:
            winners.append(excel_worksheet['B'+str(i)].value)
            losers.append(excel_worksheet['E'+str(i)].value)
        else:
            winners.append(excel_worksheet['E'+str(i)].value)
            losers.append(excel_worksheet['B'+str(i)].value)
        i += 1

    return winners, losers

######################################################
# Separate winners and losers into lists. Draw = loss.
def player_wr(excel_worksheet):

    # Prepare empty lists to populate.
    winners = []
    losers = []
    i = 1   # Iterator for rows reference.
    for row in excel_worksheet:
        # Skip ROUND separator rows.
        if "ROUND" in str(excel_worksheet['A'+str(i)].value):
            i += 1
            continue
        # Compare score, if equal (draw), both are losers.
        if excel_worksheet['C'+str(i)].value == excel_worksheet['D'+str(i)].value:
            losers.append(excel_worksheet['A'+str(i)].value)
            losers.append(excel_worksheet['F'+str(i)].value)
            i += 1
            continue
        # Check C and D columns for winner and loser, append.
        if excel_worksheet['C'+str(i)].value > excel_worksheet['D'+str(i)].value:
            winners.append(excel_worksheet['A'+str(i)].value)
            losers.append(excel_worksheet['F'+str(i)].value)
        else:
            winners.append(excel_worksheet['F'+str(i)].value)
            losers.append(excel_worksheet['A'+str(i)].value)
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
    # Sort winrate in descending order.
    winrate = dict(sorted(winrate.items(), key=lambda item:item[1], reverse=True))

    return winrate