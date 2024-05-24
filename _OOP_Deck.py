# DECLARE
class Deck:
    def __init__(self,
                 name,
                 winrate,
                 meta_share):
        self.name = name
        self.winrate = winrate
        self.metashare = meta_share

# INSTANTIATE
deck_murktide = Deck(name="Murktide",
                meta_share=25,
                winrate=50)

print(deck_murktide.name)
print(deck_murktide.winrate)