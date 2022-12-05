import pandas as pd
import numpy as np

# import seaborn as sns
# import argparse
#
# # # constructing an argument parser to parse the arguments
# # ap = argparse.ArgumentParser()
# # ap.add_argument("-f", "--file", help="path to excel file", required=True, type=str)
# # ap.add_argument("-s", "--sheet_name", help="sheet name", type=str)
# # args = vars(ap.parse_args())
# #
# # # load the excel file using argparser
# # employed_df = pd.read_excel(io=args["file"], sheet_name=args["sheet_name"])
employed_df = pd.read_excel(io="KRO.xlsx", sheet_name="Bookable")
#
# # Ashanti Region Towns in a list
ASHANTI_REGION = ["Kumasi", "Tepa", "Konongo", "Juaso", "Asokore", "Asokwa",
                  "Nkawie", "Bekwai", "Ejisu", "Ejura", "Juaben", "Mamponteng", "Kwadaso", "Mampong",
                  "Obuasi", "Offinso", "Oforikrom", "Old Tafo", "Tafo", "Suame", "Fomena", "ADANSI - FOMENA", "New Edubiase",
                  "Boaman", "Kodie", "Mankranso", "Adugyama", "Dwinyama", "Jacobu", "Edubia", "Manso Nkwanta",
                  "Agogo", "Foase", "Nyinahin", "Barekese", "Asiwa", "Kuntenase", "Kuntanase", "Tutuka", "Akomadan", "Drobonso",
                  "Nsuta", "Effiduase", "Kumawu", "Agona", "Tafo Nhyieaso", "Atafoa", "APUTUOGYA", "PRAMSO", "Bomfa",
                  "AKWATIA LINE", "Akropong", "KOKOFU", "ATWIMA KOFORIDUA", "ABUAKWA",
                  "AHENEMA KOKOBEN", "Kotwi", "Kokoben", "WIAMOASE", "BUOKROM ESTATE", "KUMASI-BO KANKYE", "SUSANSO",
                  "JUABENG", "Toase", "Trede", "ATIMATIM", "PAKYI NO.1", "PAKYI NO.2", "AHYIAEM", "ABOASO", "ONWE", "ASHTOWN",
                  "KUBEASE", "EJURA", "HWIDIEM", "Nsuta", "AGOGO-ABUAKWA", "Akrofuom", "BEKWAI ASOKWA", "FEYIASE",
                  "BOMPATA", "ANWOMASO", "DOMPOASE", "BOHYEN", "Adum", "Adum Kumasi", "MPATASIE", "ATWIMA BOKO",
                  "SEPAASE", "ANYINAM", "JUANSA", "ABOFOUR", "ADANSI", "DWENASE", "NYABO", "WIOSO", "AKUMADAN",
                  "ADANSI WIOSO", "MFENSI", "OFFINSO ABOFOUR", "MANSO AKWASISO", "NYAMEDUASE", "YAWKWEI", "ANYINASUSO",
                  "BODWESANGO", "ABUONTEM", "ATWIMA AGOGO", "WADIE ADWUMAKASE", "BEBU PAKYI NO. 2", "ADAMSO", "S/BEKWAI",
                  "SEKYEDUMASE", "AKROKERRI", "ANWIANKWANTA", "MANSO ATWERE", "EJISU BESEASE", "MANSO AMENFI",
                  "JACHIE", "TWEDIE", "YONSO", "AKRONWE", "EJISU KRAPA", "TIKROM BAWORO", "AKRAFO KOKOBEN", "KWANWOMA",
                  "AKOTOMBRA", "AHENKRO", "MAMPONGTENG", "ANKONSIA", "Dunkwa", "KONA", "ANTOAKROM", "KOFIASE", "ASONOMASO",
                  "ASANKARE", "ATOBIASE", "MANSO NKRAN", "NOBEWAM", "TWAFO WAWASE", "BANKO", "SARBIN AKROFROM", "KONONGO ODUMASI",
                  "SEPAASE TABRE", "DONASO", "BOBIN", "MAASE", "AFOSU", "SEKYEDUMASI", "MANKRASO", "FAWOADE", "NTENSERE",
                  "MOWIRE-KODIE", "OBOGU", "BONWIRE", "JAMASI", "Gyamase", "ABANKESIESO", "ODUMASE", "OFOASE",
                  "NKENKAASU", "AMANTIA", "KWAMANG", "DAMPONG", "KWABRE EAST", "DUNKWA -ON -OFFIN", "MANSO", "BEDOMASE"]

KUMASI = ["Kumasi", "Adum", "Asafo", "Ejisu", "Ayigya", "Adum Kumasi", "Asokwa", "Kwadaso", "Mamponteng", "Mampong",
           "Oforikrom", "Old Tafo", "Tafo", "Suame", "Tafo Nhyieaso", "AKWATIA LINE", "Akropong", "ABUAKWA",
            "AHENEMA KOKOBEN", "Kotwi", "Kokoben", "WIAMOASE", "BUOKROM ESTATE", "KUMASI-BO KANKYE", "SUSANSO",
          "Trede", "ATIMATIM", "PAKYI NO.1", "PAKYI NO.2", "ASHTOWN", "AGOGO-ABUAKWA", "BOHYEN", "Adum", "Adum Kumasi",
          "Nkawie", "Tanoso", "Bantama", "ASUYEBOAH NORTH", "ASUYEBOAH", "Bremang", "Adako Jachie", "EDWENEASE",
          "APRADE", "ESERESO", "PANKRONO ESTATE", "PANKRONO", "AHWIAA-KUMASI", "KRONUM ABOUHIA", "BUOKROM ASIKAFOAMBATAM",
          "NKORANSA", "NHYIAESO", "ASAFO KUMASI", "BIPAO", "APAAH", "BUOKROM", "BREMAN WEST", "AFRANCHO BRONKONG NEWSITE",
          "KUMASI CITY MALL/ASOKWA", "KENYASE", "KRONUM", "AHODWO", "AHENBRONUM", "AMPABAME", "AHINSAN", "SAWUA",
          "TARKWA MAAKRO", "DABAN", "DUASE", "MANHYIA", "ASENUA", "ADIEMMRA", "ATONSU", "OHWIM", "MAAKRO", "AHWIAA"]

BONO_REGION = ["Sunyani", "Berekum", "Dormaa Ahinkro", "Dormaa-Ahenkro", "Dormaa Ahenkro", "Drobo", "Wenchi",
               "Nsawkaw", "Sampa", "Odumasi", "Wamfie", "Banda Ahenkro", "Banda-Ahenkro", "Nkran Nkwanta",
               "Jinijini", "Atebubu", "Kintampo", "Nkoranza", "Techiman", "Kwame Danso", "Yeji", "Jema", "Busunya",
               "Tuobodom", "Kajaji", "Prang", "Goaso", "Kenyasi", "Kenyasi NO.2", "Bechem", "Hwidiem", "Kukuom",
               "Duayaw Nkwanta", "NTOTROSO", "D/NKWANTA", "DUAYAW/NKWANTA", "ABESIM", "MIM", "ACHERENSUA",
               "KOKOAA", "NSOATRE", "KENYASI NO1", "TWIMIA", "BEREKUM-SAMPA", "NSAWKAW TAIN", "SUBINSO NO2",
               "WAMANAFO", "SANKORE", "AKONKONTIWA", "NCHIRAA", "DONKRO NKWANTA", "ABOTOASE", "BEREKUM - JINIJINI",
               "YAMFO", "DERMA", "AMASU", "AFRISIPA", "KASAPIN", "NKASEIM", "ADAMSU", "KRANKA", "JAPEKROM",
               "KENYASI NO.1", "GULUMPE", "TECHIMANTIA", "ASANTEKWA", "BOMAA", "NKRANKWANTA", "BABATO-KUMA", "AMANTIN",
               "BADU", "SEIKWA", "NOBERKAW", "CHIRAA", "DEBIB"]

WESTERN_REGION = ["Enchi", "Bibiani", "Sefwi Wiawso", "Adabokrom", "Essam-Dabiso", "Essam", "Sefwi Essam", "Bodi",
                  "Juaboso", "Akontombra", "Dadieso", "SEFWI ASAWINSO", "SEFWI YAMATWA", "SEFWI YAMATWA", "SEFWI DWINASE",
                  "SEFWI BOAKO", "SEFWI CHIRANO", "SEFWI BEKWAI", "SEFWI-TANOSO", "SEFWI ASAFO", "SEFWI JUABOSO", "SEFWI AKOTI",
                  "SEFWI BOINZAN", "SEFWI AHWIAA", "SEFWI-AFERE", "SEFWI WENCHI", "SEFWI-BODI", "SEFWI BODI", "SEFWI DWENASE",
                  "SEFWI DEBISO", "sefwi koti", "SEFWI DOMEABRA", "Sekondi – Takoradi", "Takoradi", "Abuesi", "Tarkwa Nsuaem", "Tarkwa Nsuaem",
                  "Nzema East", "Sefwi Wiaso", "Bibiani–Anhwiaso Bekwai", "ADANSI ANHWIASO", "Jomoro", "Ahanta West",
                  "Amenfi East", "Prestea Huni Valley", "Shama", "Sefwi Akontombra", "Ellembele", "Wassa Amenfi Central",
                  "Wassa Amenfi West","Wassa Amenfi", "Bia West", "Bia East", "Suaman", "Aowin", "Wassa East", "Mpohor", "Juabeso",
                  "Bodie", "Daboase", "APOWA", "DAMANG", "BENSO", "BOPP", "TAKINTA", "ASANKRAGWA", "WASA DADIESO",
                  "S/AHOKWA", "ANKWASO", "ANKWAASO", "YAYASO", "WASSA AGONA", "SUHYENSO", "S/ANHWIAM", "ANHWIAM", "WASSA ASIKUMA",
                  "HIAWA", "HOTOPO", "AWUDUA", "TARKWA BONSO", "AHEBENSO", "KEJABRIL", "AGONA AMENFI", "WASA ABRESHIA",
                  "WASSA JAPA", "SAYERANO", "WASSA", "KUBI", "MFUOM", "TARKWA-BOGOSO", "AWASO", "BAWDIE", "ADJOAFUA",
                  "SAMREBOI", "ASANKRANGWA", "ABASO", "S/DWINASE", "S/WIAWSO", "WASSA AKYEMPIM", "DWINASE", "NAKABA",
                  "DOMEABRAH", "DEBISO", "KEDADWEN", "BONSU", "AMPAIN", "AHWIAA", "ADJAKAA MANSO", "ASAWINSO",
                  "YAWMATWA", "WASSA AKEMPEM", "WASSA DADIESO", "WASSA AKROPONG", "NSAWORA", "WASA KWAMANG"]

NORTHERN_REGION = ["Tamale", "Gusheigu", "Bimbilla", "Sagnerigu", "Savelugu", "Yendi", "Karaga", "Kpandai", "Kumbungu",
                   "Sang", "Nanton", "Wulensi", "Saboba", "Tatale", "Tolon", "Zabzugu", "Nalerigu", "Gambaga", "Walewale",
                   "Bunkpurugu", "Chereponi", "Yagaba", "Yunyoo", "Bole", "Salaga", "Damango", "Sawla", "Buipe", "Daboya",
                   "Kpalbe", "RICE CITY", "TISHIGU", "DALON", "GUMBUGU", "NYOLI", "BUI", "SAGNARIGU", "CHAMBA", "YAPEI",
                   "BANDA NKWANTA", "JAMA", "MPAHA", "GUSHIEGU", "SAWLA/TUNA/KALBA", "NYANKPALA", "BAMBOI", "LARIBANGA",
                   "SAVELEGU", "FUFULSO", "MAKANGO","DAMONGO"]

UPPER_EAST = ["Bawku", "Bolga", "Bolgatanga", "Navrongo", "Zebilla", "Binduri", "Zuarungu", "Bongo", "Sandema",
              "Fumbisi", "Garu", "Paga", "Nangodi", "Pusiga", "Tongo", "Tempane", "BOLGAANGA", "YIKENE",
              "PELUNGU", "YIKENE-SOUTH", "SIRIGU", "NYARIGA", "BUGRI", "BAAZUA", "WIDANA", "BAZUA", "kongo"]

UPPER_WEST = ["Wa", "Wa East", "Wa West", "Sissala East", "Sissala", "Sissala West", "Lambussie", "Jirapa", "Nandom", "Lawra"
              , "WA SOMBO", "WECHIAU", "POYENTANGA", "ULLO", "KABANYE", "GWOLLU", "LASSIA TUOLU", "BIRIFOH",
              "kulpawn", "KALEO", "HAMILE", "BABILE", "TUMU", "NADOWLI", "NANDDOM"]

NORTHERN_ZONE = {
    "KUMASI": KUMASI,
    "ASHANTI_REGION": ASHANTI_REGION,
    "WESTERN_REGION": WESTERN_REGION,
    "BONO_AHAFO_REGION": BONO_REGION,
    "NORTHERN_REGION": NORTHERN_REGION,
    "UPPER_EAST_REGION": UPPER_EAST,
    "UPPER_WEST_REGION": UPPER_WEST
}

idxs = []

copy_employed_df = employed_df
for region, towns in NORTHERN_ZONE.items():
    for town in towns:
        idx = employed_df.index[employed_df["EMP_TOWN"] == town.upper()].tolist()
        idxs.extend(idx)
        employed_df = employed_df.drop(idx)

    df = copy_employed_df.iloc[idxs]
    idxs = []

    try:
        df.to_excel(f"{region}.xlsx")
    except Exception:
        print(f"Error: {Exception}")
    else:
        print(f"{region} Files Saved")

employed_df.to_excel("KRO_Updated.xlsx")
print("DONE!!!!")





