"""Build a populated client quote from the Walvis Bay RFQ.

Source RFQ:  /root/.claude/uploads/3cb3c9ac-e77f-421c-a740-4c43c3a1baaa/f71d8301-RFQ_Provision_supply_at_Walvis_Bay.xlsx
Cost basis:  WQTEC2478.2 AURUM quote (PDF) - prices used as cost.
Markup:      35% applied to every matched line.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RFQ = "/root/.claude/uploads/3cb3c9ac-e77f-421c-a740-4c43c3a1baaa/f71d8301-RFQ_Provision_supply_at_Walvis_Bay.xlsx"
OUT = "/home/user/Visas-program/Quote_WalvisBay_marked_up_35pct.xlsx"
MARKUP = 1.35

# (RFQ description match -> (Daron code, Daron description, cost, optional remark))
# None for cost = no match found
MATCHES = {
    "Hamburger Patties":                ("21000000018", "BEEF HAMBURGER PATTIES CUT KG", 7.50, ""),
    "Ox-liver":                         ("21000000022", "BEEF LIVER WHOLE KG", 5.14, ""),
    "Silverside":                       ("21000000039", "BEEF SILVERSIDE WHOLE KG", 12.10, ""),
    "Soup bones":                       ("21000000023", "BEEF MEAT BONES KG", 2.59, ""),
    "Back bacon":                       ("2100000056",  "BACON BACK KG", 12.19, ""),
    "Streaky bacon":                    ("21000000004", "BACON STREAKY", 14.44, ""),
    "Cooked ham sliced":                ("21000000073", "COOKED HAM", 14.28, ""),
    "Mortadella":                       ("21000000100", "MORTADELLA", 6.31, ""),
    "Mortadella chicken":               ("20000001002", "CHICKEN ROLL SMOKED KG", 8.39, "approx: chicken roll smoked"),
    "Salami":                           ("20000000233", "SALAMI SPANISH KG", 8.47, ""),
    "Jagdwurst":                        ("21000000083", "JAGWURST", 5.43, ""),
    "Kabanossi":                        ("21000007008", "CABANOSSI KG", 14.70, ""),
    "Vienna sausage":                   ("21000000173", "VIENNA SMOKED KG", 6.44, ""),
    "Russians":                         ("21000000161", "RUSSIANS KG", 7.47, ""),
    "Bockwurst":                        ("21000000055", "BOCKWURST", 6.79, ""),
    "Bratwurst":                        ("21000000056", "BRATWURST", 7.41, ""),
    "Lamb legs bone-in":                ("21000000088", "LAMB LEG KG", 9.45, ""),
    "Pork belly boneless":              ("21000000122", "PORK BELLY DEBONED WHOLE KG", 10.11, ""),
    "Pork fry sausages":                ("21000000143", "PORK SAUSAGE", 6.06, ""),
    "Pork loin boneless":               ("21000000136", "PORK LOIN BONE OUT KG", 9.00, ""),
    "Pork neck boneless":               ("21000000140", "PORK NECK BONELESS KG", 11.65, ""),
    "Pork spare ribs":                  ("21000000149", "PORK SPARE RIB", 8.49, ""),
    "Liverpaste":                       ("112307",      "LIVER PASTE 200GR", 1.81, ""),
    "Luncheon meat chicken":            ("6001310009046", "LUNCHEON ROLLS CHICKEN 300G", 2.51, ""),
    "Calamares whole":                  ("1213",        "CALAMARI TUBES KG", 9.48, ""),
    "Tilapia":                          ("11143",       "ANGEL WHOLE KG", 3.30, "approx: angel fish (no tilapia listed)"),
    "Hake fillet":                      ("1208",        "HAKE FILLETS 4-6 KG", 9.94, ""),
    "Salmon":                           ("221",         "SALMON GUTTED NORWEGIAN KG", 20.67, ""),
    "Sardines in oil":                  ("520732",      "SARDINES IN OIL SAUCE 155g", 1.53, ""),
    "Tuna in oil":                      ("6001115001368", "TUNA SHREDDED V/OIL LUCKY STAR 170G", 1.96, ""),
    "Oysters":                          ("13701080",    "OYSTER FRESH EACH", 1.35, "fresh each (RFQ asked TIN)"),
    "Chicken burger":                   ("789654123",   "CHICKEN SCHNITZEL CRUMB KG", 5.74, "approx: schnitzel crumb"),
    "Chicken breast":                   ("203",         "CHICKEN BREAST FILLETS KG", 10.78, ""),
    "Chicken legs":                     ("201",         "CHICKEN QUARTER LEGS KG", 6.80, ""),
    "Apples green":                     ("402",         "APPLES G.S LS KG", 3.17, ""),
    "Apples red":                       ("403",         "APPLES RED TOP LS KG", 3.15, ""),
    "Bananas":                          ("405",         "BANANAS KG", 2.64, ""),
    "Grapes black":                     ("407",         "GRAPES BLACK KG", 6.32, ""),
    "Grapes red":                       ("409",         "GRAPES RED SLESS KG", 6.32, ""),
    "Kiwi":                             ("411",         "KIWI FRUIT GREEN KG", 10.18, ""),
    "Lemons":                           ("412",         "LEMONS KG", 2.71, ""),
    "Oranges":                          ("420",         "ORANGES KG", 2.74, ""),
    "Papaya":                           ("421",         "PAW PAW KG", 2.67, "papaya = paw paw"),
    "Paw Paw":                          ("421",         "PAW PAW KG", 2.67, ""),
    "Pears":                            ("1541321",     "FRESH PEARS KG", 2.71, ""),
    "Pineapple":                        ("427",         "PINEAPPLE QUEEN KG", 2.41, ""),
    "Frozen mango in cubes":            ("600013",      "FROZEN MANGO PIECES IQF 1KG", 11.32, ""),
    "Lychees":                          (None, None, None, "NO MATCH - lychees not on Aurum list"),
    "Peaches":                          ("673",         "PEACH HALVES 410G", 1.89, ""),
    "Pineapple sliced":                 ("6009684163104", "PINEAPPLE RINGS 425G", 2.36, ""),
    "Mango":                            (None, None, None, "NO MATCH - tinned mango not on Aurum list"),
    "Beetroot":                         ("434",         "BEETROOT KG", 1.39, ""),
    "Cabbage, white":                   ("441",         "CABBAGE KG", 2.26, ""),
    "Carrots":                          ("445",         "CARROTS KG", 1.58, ""),
    "Cucumbers":                        ("450",         "CUCUMBERS ENGLISH KG", 5.27, ""),
    "Dill fresh":                       (None, None, None, "NO MATCH - fresh dill not listed"),
    "Eggplant":                         ("435",         "BRINJALS KG", 3.61, "brinjal = eggplant"),
    "Marrows (zucchini)":               ("432",         "BABY MARROW KG", 7.22, ""),
    "Mushrooms":                        ("459",         "FRESH MUSHROOM WHITE KG", 12.44, ""),
    "Onions":                           ("460",         "ONION KG", 1.81, ""),
    "Pumpkin":                          ("468",         "PUMPKIN WHITE KG", 0.90, ""),
    "Butternut squash":                 ("438",         "BUTTERNUT KG", 1.67, ""),
    "Red bell pepper":                  ("465",         "RED PEPPERS KG", 8.06, ""),
    "Springonions":                     ("472",         "SPRING ONION KG", 16.53, ""),
    "Tomatoes, half ripe":              ("476",         "TOMATOES KG", 2.82, ""),
    "Potatoes":                         ("467",         "POTATOES KG", 1.52, ""),
    "Frozen French fries":              ("12301",       "FRENCH FRIES 10mm 2.5kg (per kg)", 6.86, ""),
    "Frozen potato wedges":             ("CHKLQR028",   "POTATO WEDGES 4X2.5KG (per kg)", 5.06, ""),
    "Baked beans in tomato sauce":      ("6001059987520", "BAKED BEANS IN TOMATO SAUCE KOO 400G", 1.33, ""),
    "Beetroots":                        ("6001024037466", "BEETROOT SLICED KOO 780G", 3.15, ""),
    "Mushrooms sliced":                 ("6009711384397", "MUSHROOM SLICED 400G", 3.18, ""),
    "Mushrooms whole":                  ("6009684163051", "MUSHROOM WHOLE 400G", 3.20, ""),
    "Olives green":                     ("1268",        "OLIVES GREEN 200G", 1.84, ""),
    "Sweet corn":                       ("6001024386655", "CORN WHOLE KERNEL KOO 410G", 2.14, ""),
    "Tomato paste concentrated":        ("6001012310125", "TOMATO PASTE CUP 115G", 1.35, ""),
    "Tomato peeled":                    ("691",         "TOMATO WHOLE PEELED 400G", 1.77, ""),
    "Winesauerkraut":                   ("18691",       "SAUERKRAUT WINE KUEHNE 720ML", 4.58, ""),
    "Pickled gherkins whole":           ("6001024542020", "GHERKINS SWEET N SOUR 780G", 7.07, ""),
    "Pesto green":                      ("900355115410544", "FROZEN PESTO BASIL 1KG", 18.97, "frozen 1kg (RFQ asked JAR)"),
    "Red beans":                        ("6001168010515", "KIDNEY BEANS RED DRY 500G", 4.20, ""),
    "Yellow splitpeas":                 ("1242",        "DALL CHANA KG (yellow split peas)", 2.17, ""),
    "Dried shiitake mushrooms":         (None, None, None, "NO MATCH - dried shiitake not listed"),
    "Dried dill tips":                  (None, None, None, "NO MATCH - dried dill not listed"),
    "Edam":                             ("1220",        "CHEESE EDAM LOAVES KG", 15.85, ""),
    "Gouda":                            ("501",         "CHEESE GOUDA LOAVES KG", 14.13, ""),
    "Parmesan whole":                   ("PAR34001",    "PARMESAN CHEESE ROUNDS", 29.81, ""),
    "Cream cheese":                     ("064621",      "CHEESE CREAM 230g", 4.48, ""),
    "Cottage cheese":                   ("06462",       "CHEESE COTTAGE 250ML", 3.93, ""),
    "Cheddar":                          ("503",         "CHEESE CHEDDAR YELLOW LOAVES KG", 14.70, ""),
    "Cream cooking":                    ("12171",       "CREAM VERSATIE LT", 7.08, ""),
    "Sour cream":                       ("12811",       "SOUR CREAM 2.5LTR", 34.32, "2.5L tub (RFQ asked BCK)"),
    "Eggs 30'S":                        ("202",         "EGGS TRAY OF 30", 10.47, ""),
    "Icecream chocolat":                ("6460",        "ICE CREAM TUB CHOCOLATE 2L", 8.40, ""),
    "Icecream malaga rum raisins":      ("6456",        "ICE CREAM TUB RUM N RAISIN 2L", 8.40, ""),
    "Icecream mokka-coffee":            (None, None, None, "NO MATCH - mokka/coffee ice cream not listed"),
    "Icecream strawberry":              ("6459",        "ICE CREAM TUB STRAWBERRY 2L", 8.40, ""),
    "Icecream vanille":                 ("6451",        "ICE CREAM TUB VANILLA 2L", 8.40, ""),
    "Kefir":                            ("9955995",     "OSHIKANDELA PLAIN DRINKING YOGHURT 500ML", 1.59, "approx: drinking yoghurt (no kefir listed)"),
    "Full milk UHT":                    ("6001415001600", "MILK FULL CREAM LONG LIFE 1LTR", 2.21, ""),
    "Sweet condensed milk":             ("677",         "MILK CONDENSED NESTLE 385G", 4.15, ""),
    "Fruityoghurt":                     ("110265",      "YOGHURT ASSORTED 100ML", 1.05, ""),
    "Yoghurt plain":                    ("1102655",     "YOGHURT PLAIN 100ML", 0.49, ""),
    "Olive oil extra virgin":           ("8018440003811", "OLIVE OIL EXTRA VIRGIN 1LTR", 13.97, ""),
    "Sunflower oil":                    ("6009629840046", "SUNFLOWER COOKING OIL 5L", 20.87, ""),
    "Brown bread":                      ("1302",        "BREAD BROWN", 1.13, ""),
    "Hot dog rolls":                    ("1306",        "HOTDOG ROLLS", 0.29, ""),
    "White bread":                      ("1301",        "BREAD WHITE", 1.13, ""),
    "rye bread":                        ("13070",       "RYE BREAD", 2.62, ""),
    "Paratha plain":                    ("211",         "WHOLE GRAIN TORTILLA WRAPS 20cm 8's", 3.30, "approx: tortilla wraps (no paratha listed)"),
    "Hamburgerrolls":                   ("1304",        "HAMBURGER BUNS", 0.33, ""),
    "Puff pastry":                      ("256151",      "PUFF PASTRY ROLL 400G", 1.73, ""),
    "Coconut biscuits":                 ("485478978",   "BISCUITS ROMANY CREAMS 200G", 3.14, "approx: romany creams"),
    "Cream crackers":                   ("660",         "CRACKER CREAM 200G", 1.53, ""),
    "Salticraz":                        ("656",         "CRACKER SALTICRAX 200G", 2.17, ""),
    "Chocolate chip cookies":           ("70501",       "BISCUITS CHOC CHIP HENRO 160G", 1.96, ""),
    "Digestive, chocolate":             ("60017",       "BISCUITS BETTA DIGESTIVE BAKERS 200G", 3.22, ""),
    "Ginger nuts biscuits":             ("70502",       "BISCUITS GINGER HENRO 175G", 1.96, ""),
    "Oreo cookies":                     ("02850201456520", "BISCUITS OREO 128.8G", 1.85, ""),
    "Cake mix chocolate":               ("6891",        "CAKE MIX CHOCOLATE 800G SNOWFLAKE", 3.87, ""),
    "Cake mix vanilla":                 ("6892",        "CAKE MIX VANILLA 800G SNOWFLAKE", 3.37, ""),
    "Flour wheat white":                ("564155",      "WHITE BREAD WHEAT FLOUR 1KG", 1.93, ""),
    "Kellogg's Rice Crispies":          ("112319",      "RICE CRISPIES KELLOGS 340G", 4.24, ""),
    "Nutella choco spread":             ("664",         "SPREAD NUTELLA 350GR", 8.51, ""),
    "Honey":                            ("699",         "HONEY PURE SQUEEZY BOTTLE 500GR", 6.04, ""),
    "Peanut butter":                    ("665256101",   "PEANUT BUTTER SMOOTH YUM YUM 400G", 3.75, ""),
    "Sugar granulated 10 kg":           ("6008079005180", "SUGAR WHITE 10KG", 17.68, ""),
    "Coffee grounded":                  ("220611",      "COFFEE GROUND TWO BEARDS KG", 30.30, "priced per KG (RFQ asked PCK)"),
    "Nescafe gold":                     ("9561",        "COFFEE GOLD NESCAFE 200G", 13.89, ""),
    "Breakfast tea 20's":               ("6663",        "TEA ENGLISH BREAKFAST 20'S TEEKANNE", 2.93, ""),
    "Tea Camomile 20's":                ("6662",        "TEA CHAMOMILE 20's TEEKANNE", 3.42, ""),
    "Earl grey tea 20's":               ("261215120202152", "TEA EARL GREY 20's TEEKANNE", 2.93, ""),
    "Beef bouillon 12 x 10 gr tablets": ("6001038062453", "CUBES BEEF 12's 120g KNORROX", 1.26, ""),
    "Chicken bouillon powder":          ("1344",        "STOCK CHICKEN 1KG", 6.54, ""),
    "Fish bouillon":                    ("515612",      "STOCK FISH", 3.16, ""),
    "Mushroom cremesoup":               ("6001087000475", "SOUP CREAM OF MUSHROOM 1KG", 5.65, ""),
    "Onionsoup":                        ("661",         "SOUP BROWN ONION 1KG", 5.18, ""),
    "Barbecue sauce":                   ("6009699177034", "BBQ SAUCE LAPPIES 500ML", 3.01, ""),
    "Chili sauce":                      ("56812511",    "CHILLI SAUCE SHRIRACHA 435ml", 5.37, ""),
    "Chili garlic sauce":               ("5565",        "SAUCE CHILLI GARLIC 300ml", 3.25, ""),
    "Sweet chilisauce":                 ("600102001501675", "SAUCE SWEET CHILLI WELLINGTONS 700ML", 3.18, ""),
    "Fishsauce":                        ("679",         "SAUCE FISH CA-LINH 690ML", 2.60, ""),
    "Mango chutney":                    ("6001092500052.1", "CHUTNEY ASSORTED MRS BALLS 860G", 3.10, ""),
    "Mayonnaise":                       ("6009522301262", "MAYONNAISE CROSSE & BLACKWELL TANGY 750G", 3.18, ""),
    "Mayonnaise Hellmans squeeze bottle": ("N5641541",  "NOLA MAYONNAISE SQUEEZE", 4.08, "approx: Nola squeeze (no Hellmans listed)"),
    "Thousand Island dressing":         ("35456",       "SALAD DRESSING 1000 ISLANDS KNORR 340ML", 3.17, ""),
    "Tomato ketchup Heinz":             ("60019578",    "TOMATO SAUCE ALL GOLD 700ML", 3.02, "approx: All Gold (no Heinz listed)"),
    "Sweet chili sauce":                ("600102001501675", "SAUCE SWEET CHILLI WELLINGTONS 700ML", 3.18, ""),
    "Pasta sauce":                      ("5464651",     "ALL GOLD PASTA SAUCE 405G", 3.93, ""),
    "Pizza sauce":                      ("6001168010614", "PIZZA SAUCE 4.1kg", 17.61, "4.1kg pack"),
    "Tartar sauce":                     ("45611",       "TARTAR SAUCE 2L", 19.44, "2L pack"),
    "Taco sauce":                       (None, None, None, "NO MATCH - taco sauce not listed"),
    "Teriyaki sauce":                   ("651531200",   "TERIYAKI SAUCE 295ML", 4.75, ""),
    "Tom yum paste":                    ("235987",      "TOM YUM PASTE 454g SUREE", 7.07, ""),
    "Curry powder":                     ("752020",      "CURRY POWDER KG", 7.33, ""),
    "Garlic powder":                    ("39534766",    "GARLIC POWDER 1KG", 10.66, ""),
    "Mixed herbs":                      ("1654377",     "ROBERTSONS MIXED HERBS 100ML", 2.64, ""),
    "Oregano crushed":                  ("6631000",     "OREGANUM REFILL 100G", 1.26, ""),
    "Paprika powder sweet":             ("1337",        "PAPRIKA GROUND 1KG", 7.23, ""),
    "Pepper black grounded":            ("AL113",       "PEPPER BLACK GROUND KG", 23.66, ""),
    "Chicken spices":                   ("6001087307604", "CHICKEN SPICE 1 KG", 4.24, ""),
    "Onion powder":                     ("6001038009359", "ONION POWDER 1KG", 7.95, ""),
    "Kitchensalt":                      ("19657802",    "COARSE SALT 2KG", 0.63, "per 2kg pack"),
    "Table salt iodised flask":         ("6008145000026", "SALT TABLE BAG FINE 1KG", 0.57, ""),
    "Tandoori masala":                  ("1137318",     "TANDOORI MASALA 1kg", 11.00, ""),
    "Chana masala":                     ("1242001",     "CHANA MASALA 100g", 3.61, ""),
    "Mutton masala":                    ("1102131",     "MASALA MUTTON 100g", 3.61, ""),
    "Chicken masala":                   ("9918",        "MASALA CHICKEN 100g", 5.42, ""),
    "Biryani masala":                   ("1239",        "MASALA BIRYANI 100g", 5.42, ""),
    "Fish curry masala":                (None, None, None, "NO MATCH - fish curry masala not listed"),
    "Phav baji masala":                 ("1137318",     "MASALA PAV BHAJI 100g", 3.61, ""),
    "Macaroni penne":                   ("1973",        "PASTA PENNE BARILLA/RIGATE 500G", 3.32, ""),
    "Macaroni elbows":                  ("22112233",    "PASTA ELBOW POLANA 500G", 1.16, ""),
    "Penne Rigate Barilla":             ("1973",        "PASTA PENNE BARILLA/RIGATE 500G", 3.32, ""),
    "Spaghetti":                        ("60095416544", "PASTA SPAGHETTI POLANA 500G", 2.02, ""),
    "Fettuccine":                       ("RIG10005765", "PASTA LINGUINI 500G", 1.77, "approx: linguini (no fettuccine)"),
    "Cup noodles beef":                 ("1119",        "NOODLES INSTANT CUP 70G BEEF", 0.89, ""),
    "Cup noodles chicken":              ("1119",        "NOODLES INSTANT CUP 70G CHICKEN", 0.89, ""),
    "Cup noodles seafood":              ("1119",        "NOODLES INSTANT CUP 70G SEAFOOD", 0.89, ""),
    "Basmati rice":                     ("4870457860030", "RICE BASMATI WHITE ORIENT (per kg)", 3.07, "5kg bag $15.35 = $3.07/kg"),
    "Apple juice":                      ("6009615146641", "JUICE FRUIT 100% APPLE 1L", 2.37, ""),
    "Grape (raisins) juice":            ("6009615146644", "JUICE FRUIT 100% RED GRAPE 1L", 2.37, ""),
    "Grapefruit juice":                 (None, None, None, "NO MATCH - grapefruit juice not listed (only red grape)"),
    "Mango juice":                      ("6009615146643", "JUICE FRUIT 100% MANGO 1L", 2.37, ""),
    "Orange juice":                     ("6009615146642", "JUICE FRUIT 100% ORANGE 1L", 2.37, ""),
    "Pineapple juice":                  ("6009615146645", "JUICE FRUIT 100% PINEAPPLE 1L", 2.37, ""),
    "Tomato juice":                     ("1126",        "JUICE TOMATO LTR", 3.30, ""),
    "Tropical juice":                   ("6009615146646", "JUICE FRUIT 100% TROPICAL 1L", 2.37, ""),
    "Peach  juice":                     ("6009615146649", "JUICE FRUIT 100% PEACH 1L", 2.37, ""),
    "Guava juice":                      ("6009615146647", "JUICE FRUIT 100% GUAVA 1L", 2.37, ""),
    "Still water PET 6x1,5 L":          ("3154",        "WATER STILL 12x1.5L MINAQUA", 5.87, "Daron pack 12x1.5L (RFQ asked 6x)"),
    "Cooking wine red":                 ("6002323400233", "WINE RED 750M", 9.62, ""),
}


def lookup(desc):
    key = desc.strip()
    if key in MATCHES:
        return MATCHES[key]
    # try case-insensitive trimmed
    for k, v in MATCHES.items():
        if k.lower() == key.lower():
            return v
    return (None, None, None, "NO MATCH")


def main():
    src = openpyxl.load_workbook(RFQ, data_only=True)
    src_ws = src["Sheet1"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quote"

    headers = [
        "Description (RFQ)",
        "Qty",
        "Unit",
        "Daron Code",
        "Daron Description",
        "Cost USD",
        "Markup %",
        "Unit Price USD (Cost x 1.35)",
        "Line Total USD",
        "Remarks",
    ]
    bold = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="305496")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = bold
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    out_row = 2
    subtotal = 0.0
    matched_count = 0
    unmatched_count = 0

    rows = list(src_ws.iter_rows(values_only=True))
    for r in rows[1:]:  # skip header
        desc, qty, unit = r[0], r[1], r[2]
        if not desc:
            continue
        code, dname, cost, remark = lookup(desc)
        ws.cell(row=out_row, column=1, value=desc.strip() if isinstance(desc, str) else desc)
        ws.cell(row=out_row, column=2, value=qty)
        ws.cell(row=out_row, column=3, value=unit)
        ws.cell(row=out_row, column=4, value=code or "")
        ws.cell(row=out_row, column=5, value=dname or "")

        if cost is not None:
            unit_price = round(cost * MARKUP, 2)
            line_total = round(unit_price * (qty or 0), 2)
            ws.cell(row=out_row, column=6, value=cost).number_format = '"$"#,##0.00'
            ws.cell(row=out_row, column=7, value="35%")
            ws.cell(row=out_row, column=8, value=unit_price).number_format = '"$"#,##0.00'
            ws.cell(row=out_row, column=9, value=line_total).number_format = '"$"#,##0.00'
            subtotal += line_total
            matched_count += 1
        else:
            ws.cell(row=out_row, column=7, value="35%")
            unmatched_count += 1
            unmatched_fill = PatternFill("solid", fgColor="FFE699")
            for col in range(1, 11):
                ws.cell(row=out_row, column=col).fill = unmatched_fill

        ws.cell(row=out_row, column=10, value=remark)

        for col in range(1, 11):
            ws.cell(row=out_row, column=col).border = border
            ws.cell(row=out_row, column=col).alignment = Alignment(vertical="center", wrap_text=True)
        out_row += 1

    # totals block
    out_row += 1
    label_font = Font(bold=True)
    ws.cell(row=out_row, column=8, value="SUBTOTAL (USD)").font = label_font
    sub_cell = ws.cell(row=out_row, column=9, value=round(subtotal, 2))
    sub_cell.number_format = '"$"#,##0.00'
    sub_cell.font = label_font
    out_row += 1
    ws.cell(row=out_row, column=8, value=f"Matched lines: {matched_count}")
    ws.cell(row=out_row, column=9, value=f"Unmatched lines: {unmatched_count}")

    # Column widths
    widths = [38, 8, 8, 18, 38, 11, 10, 14, 14, 50]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 32

    # Title block above (insert rows)
    ws.insert_rows(1, amount=4)
    title = ws.cell(row=1, column=1, value="QUOTE - Provision Supply at Walvis Bay (35% markup applied to Aurum cost basis)")
    title.font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    ws.cell(row=2, column=1, value="Cost basis: Daron quote WQTEC2478.2 (AURUM, 09/05/2026)")
    ws.cell(row=3, column=1, value="Markup: 35% on each line (Unit Price = Cost x 1.35)")
    ws.cell(row=4, column=1, value="Highlighted yellow rows = no match in cost source; unit price needed.")

    wb.save(OUT)
    print(f"Wrote: {OUT}")
    print(f"Matched: {matched_count}   Unmatched: {unmatched_count}   Subtotal: ${subtotal:,.2f}")


if __name__ == "__main__":
    main()
