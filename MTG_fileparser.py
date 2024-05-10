# Decklist file name
DECK_LIST = "Deck.txt"
# Collection file name
EXCEL_FILE = "Collection.xlsx"

from collections import Counter

#############################################################################
# Method returns fully counted [COUNT, CARD NAME] list, given a txt decklist.
def decklist_parse():

    # Load decklist txt file.
    deck = open("./" + DECK_LIST)

    # Deck -> .txt to Python list
    decklist = []
    for row in deck.readlines():
        if row.find(" ") == -1: decklist.append("\n")
        if row.find(" ") > 0:
            if row[row.find(" ")+1:row.find("\n")] in decklist:
                decklist[decklist.index(row[row.find(" ")+1:row.find("\n")])][1] += int(row[:row.find(" ")])
            try: decklist.append([int(row[:row.find(" ")]), row[row.find(" ")+1:row.find("\n")]])
            except: continue

    # Deck -> Alphabetized pile
    pile = []
    i = 0
    for element in decklist:
        if element != "\n":
            j = int(decklist[i][0])
            for k in range(j): pile.append(decklist[i][1])
        i += 1
    pile.sort()

    # Deck -> Ordered and counted list
    c_pile = Counter(pile)
    counted_list = []
    for key in c_pile.keys(): counted_list.append([int(c_pile[key]), key])

    return(counted_list)

##########################################################################
# Method returns [COUNT, CARD NAME, LOCATION], given a card name argument.
def collection_check(card):

    # Load collection Excel file.
    import openpyxl
    worksheet = openpyxl.load_workbook("./" + EXCEL_FILE).active

    # Populate the mtg_collection list.
    mtg_collection = [] # [CARD NAME (col A), LOCATION (col E)]
    i = 2 # Start with row 2.
    for row in worksheet:
        row = [str(worksheet['A'+str(i)].value), str(worksheet['E'+str(i)].value)]
        if row != ['None', 'None']:
            mtg_collection.append(row)
        i += 1

    # Store each instance of a card in a temp list...
    loc_temp = [] # [LOCATION, ..., LOCATION]
    i = 0
    for row in mtg_collection:
        if card in row:
            loc_temp.append(row[1])
            i += 1
    # ...then transform it to the loc_count dict.
    loc_count = Counter(loc_temp)

    # Populate the cards_in_collection list of lists.
    cards_in_collection = [] # [[COUNT, CARD NAME, LOCATION], ...]
    for key in loc_count.keys():
        cards_in_collection.append([int(loc_count[key]), card, key])

    return(cards_in_collection)

##################################################
# Method returns a single float of a card's price.
def get_price(card):

    # Scryfall API HTTP GET request for card name.
    import requests
    url = "https://api.scryfall.com/cards/search?q=%22" + card + "%22"

    # Extraction of card price in EUR (CardMarket).
    response = requests.get(url).json()
    r_data = response.get("data")
    try: data_to_dict = r_data[0]
    except: return(0)
    r_prices = data_to_dict.get("prices")
    try: price_eur = float(r_prices.get("eur"))
    except: return(0)

    return(price_eur)