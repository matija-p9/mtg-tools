from deck_collection import deck_functions as df

# Compare the decklist with the collection, write down what's missing.
decklist = df.decklist_parse()
missing = []
# Iterate through decklist.
for line in decklist:
    coll_check = df.collection_check(line[1])
    coll_card_sum = 0
    # Iterate through collection, append to missing.
    for element in coll_check:
        coll_card_sum += element[0]
    if coll_check == []:
        missing.append([line[0], line[1]])
    elif coll_card_sum < line[0]:
        missing.append([line[0]-coll_card_sum, line[1]])

# List the missing cards, with prices and total sum, from Scryfall.
sum_total = 0
for element in missing:
    element.insert(2, df.get_price(element[1]))
    element.insert(3, element[0]*element[2])
    sum_total += element[0]*element[2]
    print(element[0] # Card Count
         ,element[1] # Card Name
         ,"%.2f" % element[2] # Card Price
         ,"%.2f" % element[3] # Card Count Price
         )
print("%.2f" % sum_total) # Total Price