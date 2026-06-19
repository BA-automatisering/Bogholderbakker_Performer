from datetime import date

manuelliste = []
driftliste = []
aktuel_bogholderbakke = ""
item_count = 0

max_retry_count = 3
if globals.aktuel_bogholderbakke == "Fakturabeslut.08: Håndter afvist faktura":
    globals.max_retry_count=5
    
start = date.today()
slut = date.today()
