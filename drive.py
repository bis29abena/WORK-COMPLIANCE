import pandas as pd
import numpy as np

# import seaborn as sns
# import argparse
#
# # # constructing an argument parser to parse the arguments
# # ap = argparse.ArgumentParser()
# # ap.add_argument('  f', '    file', help='path to excel file', required=True, type=str)
# # ap.add_argument('  s', '    sheet_name', help='sheet name', type=str)
# # args = vars(ap.parse_args())
# #
# # # load the excel file using argparser
# # employed_df = pd.read_excel(io=args['file'], sheet_name=args['sheet_name'])
employed_df = pd.read_excel(io='Data.xlsx', sheet_name='new')
#
# # Ashanti Region Towns in a list
ASHANTI_REGION = ['BEHENASE', 'Bosomtwe', 'EFFIDUASI-ASH', 'Kumasi', 'Tepa', 'Konongo', 'Juaso', 'Asokore', 'Asokwa', 'ASHANTI',
                  'Nkawie', 'Bekwai', 'Ejisu', 'Ejura', 'Juaben', 'Mamponteng', 'Kwadaso', 'Mampong',
                  'Obuasi', 'Offinso', 'Oforikrom', 'Old Tafo', 'Tafo', 'Suame', 'Fomena', 'ADANSI  FOMENA', 'New Edubiase',
                  'Boaman', 'Kodie', 'Mankranso', 'Adugyama', 'Dwinyama', 'Jacobu', 'Edubia', 'Manso Nkwanta',
                  'Agogo', 'Foase', 'Nyinahin', 'Barekese', 'Asiwa', 'Kuntenase', 'Kuntanase', 'Tutuka', 'Akomadan', 'Drobonso',
                  'Nsuta', 'Effiduase', 'Kumawu', 'Agona', 'Tafo Nhyieaso', 'Atafoa', 'APUTUOGYA', 'PRAMSO', 'Bomfa',
                  'AKWATIA LINE', 'Akropong', 'KOKOFU', 'ATWIMA KOFORIDUA', 'ABUAKWA',
                  'AHENEMA KOKOBEN', 'Kotwi', 'Kokoben', 'WIAMOASE', 'BUOKROM ESTATE', 'KUMASI  BO KANKYE', 'SUSANSO',
                  'JUABENG', 'Toase', 'Trede', 'ATIMATIM', 'PAKYI NO.1', 'PAKYI NO.2', 'AHYIAEM', 'ABOASO', 'ONWE', 'ASHTOWN',
                  'KUBEASE', 'EJURA', 'HWIDIEM', 'Nsuta', 'AGOGO  ABUAKWA', 'Akrofuom', 'BEKWAI ASOKWA', 'FEYIASE',
                  'BOMPATA', 'ANWOMASO', 'DOMPOASE', 'BOHYEN', 'Adum', 'Adum Kumasi', 'MPATASIE', 'ATWIMA BOKO',
                  'SEPAASE', 'ANYINAM', 'JUANSA', 'ABOFOUR', 'ADANSI', 'DWENASE', 'NYABO', 'WIOSO', 'AKUMADAN',
                  'ADANSI WIOSO', 'MFENSI', 'OFFINSO ABOFOUR', 'MANSO AKWASISO', 'NYAMEDUASE', 'YAWKWEI', 'ANYINASUSO',
                  'BODWESANGO', 'ABUONTEM', 'ATWIMA AGOGO', 'WADIE ADWUMAKASE', 'BEBU PAKYI NO. 2', 'ADAMSO', 'S/BEKWAI',
                  'SEKYEDUMASE', 'AKROKERRI', 'ANWIANKWANTA', 'MANSO ATWERE', 'EJISU BESEASE', 'MANSO AMENFI',
                  'JACHIE', 'TWEDIE', 'YONSO', 'AKRONWE', 'EJISU KRAPA', 'TIKROM BAWORO', 'AKRAFO KOKOBEN', 'KWANWOMA',
                  'AKOTOMBRA', 'AHENKRO', 'MAMPONGTENG', 'ANKONSIA', 'Dunkwa', 'KONA', 'ANTOAKROM', 'KOFIASE', 'ASONOMASO',
                  'ASANKARE', 'ATOBIASE', 'MANSO NKRAN', 'NOBEWAM', 'TWAFO WAWASE', 'BANKO', 'SARBIN AKROFROM', 'KONONGO ODUMASI',
                  'SEPAASE TABRE', 'DONASO', 'BOBIN', 'MAASE', 'AFOSU', 'SEKYEDUMASI', 'MANKRASO', 'FAWOADE', 'NTENSERE',
                  'MOWIRE  KODIE', 'OBOGU', 'BONWIRE', 'JAMASI', 'Gyamase', 'ABANKESIESO', 'ODUMASE', 'OFOASE',
                  'NKENKAASU', 'AMANTIA', 'KWAMANG', 'DAMPONG', 'KWABRE EAST', 'DUNKWA   ON   OFFIN', 'MANSO', 'BEDOMASE']

KUMASI = ['Kumasi', 'Adum', 'Asafo', 'Ejisu', 'Ayigya', 'Adum Kumasi', 'Asokwa', 'Kwadaso', 'Mamponteng', 'Mampong',
          'Oforikrom', 'Old Tafo', 'Tafo', 'Suame', 'Tafo Nhyieaso', 'AKWATIA LINE', 'Akropong', 'ABUAKWA',
          'AHENEMA KOKOBEN', 'Kotwi', 'Kokoben', 'WIAMOASE', 'BUOKROM ESTATE', 'KUMASI  BO KANKYE', 'SUSANSO',
          'Trede', 'ATIMATIM', 'PAKYI NO.1', 'PAKYI NO.2', 'ASHTOWN', 'AGOGO  ABUAKWA', 'BOHYEN', 'Adum', 'Adum Kumasi',
          'Nkawie', 'Tanoso', 'Bantama', 'ASUYEBOAH NORTH', 'ASUYEBOAH', 'Bremang', 'Adako Jachie', 'EDWENEASE',
          'APRADE', 'ESERESO', 'PANKRONO ESTATE', 'PANKRONO', 'AHWIAA  KUMASI', 'KRONUM ABOUHIA', 'BUOKROM ASIKAFOAMBATAM',
          'NKORANSA', 'NHYIAESO', 'ASAFO KUMASI', 'BIPAO', 'APAAH', 'BUOKROM', 'BREMAN WEST', 'AFRANCHO BRONKONG NEWSITE',
          'KUMASI CITY MALL/ASOKWA', 'KENYASE', 'KRONUM', 'AHODWO', 'AHENBRONUM', 'AMPABAME', 'AHINSAN', 'SAWUA',
          'TARKWA MAAKRO', 'DABAN', 'DUASE', 'MANHYIA', 'ASENUA', 'ADIEMMRA', 'ATONSU', 'OHWIM', 'MAAKRO', 'AHWIAA']

BONO_REGION = ['Sunyani', 'Berekum', 'Dormaa Ahinkro', 'Dormaa  Ahenkro', 'Dormaa Ahenkro', 'Drobo', 'Wenchi', 'Dormaa'
               'Nsawkaw', 'Sampa', 'Odumasi', 'Wamfie', 'Banda Ahenkro', 'Banda  Ahenkro', 'Nkran Nkwanta',
               'Jinijini', 'Atebubu', 'Kintampo', 'Nkoranza', 'Techiman', 'Kwame Danso', 'Yeji', 'Jema', 'Busunya',
               'Tuobodom', 'Kajaji', 'Prang', 'Goaso', 'Kenyasi', 'Kenyasi NO.2', 'Bechem', 'Hwidiem', 'Kukuom',
               'Duayaw Nkwanta', 'NTOTROSO', 'D/NKWANTA', 'DUAYAW NKWANTA', 'ABESIM', 'MIM', 'ACHERENSUA',
               'KOKOAA', 'NSOATRE', 'KENYASI NO1', 'TWIMIA', 'BEREKUM  SAMPA', 'NSAWKAW TAIN', 'SUBINSO NO2',
               'WAMANAFO', 'SANKORE', 'AKONKONTIWA', 'NCHIRAA', 'DONKRO NKWANTA', 'ABOTOASE', 'BEREKUM  JINIJINI',
               'YAMFO', 'DERMA', 'AMASU', 'AFRISIPA', 'KASAPIN', 'NKASEIM', 'ADAMSU', 'KRANKA', 'JAPEKROM',
               'KENYASI NO.1', 'GULUMPE', 'TECHIMANTIA', 'ASANTEKWA', 'BOMAA', 'NKRANKWANTA', 'BABATO  KUMA', 'AMANTIN',
               'BADU', 'SEIKWA', 'NOBERKAW', 'CHIRAA', 'DEBIB']

WESTERN_REGION = ['Enchi', 'Bibiani', 'Sefwi Wiawso', 'Adabokrom', 'Essam  Dabiso', 'Essam', 'Sefwi Essam', 'Bodi',
                  'Juaboso', 'Akontombra', 'Dadieso', 'SEFWI ASAWINSO', 'SEFWI YAMATWA', 'SEFWI YAMATWA', 'SEFWI DWINASE',
                  'SEFWI BOAKO', 'SEFWI CHIRANO', 'SEFWI BEKWAI', 'SEFWI  TANOSO', 'SEFWI ASAFO', 'SEFWI JUABOSO', 'SEFWI AKOTI',
                  'SEFWI BOINZAN', 'SEFWI AHWIAA', 'SEFWI  AFERE', 'SEFWI WENCHI', 'SEFWI  BODI', 'SEFWI BODI', 'SEFWI DWENASE',
                  'SEFWI DEBISO', 'sefwi koti', 'SEFWI DOMEABRA', 'Sekondi Takoradi', 'Takoradi', 'Abuesi', 'Tarkwa Nsuaem', 'Tarkwa Nsuaem', 'sekondi',
                  'Nzema East', 'Sefwi Wiaso', 'Bibiani Anhwiaso Bekwai', 'ADANSI ANHWIASO', 'Jomoro', 'Ahanta West',
                  'Amenfi East', 'Prestea Huni Valley', 'Shama', 'Sefwi Akontombra', 'Ellembele', 'Wassa Amenfi Central',
                  'Wassa Amenfi West', 'Wassa Amenfi', 'Bia West', 'Bia East', 'Suaman', 'Aowin', 'Wassa East', 'Mpohor', 'Juabeso',
                  'Bodie', 'Daboase', 'APOWA', 'DAMANG', 'BENSO', 'BOPP', 'TAKINTA', 'ASANKRAGWA', 'WASA DADIESO',
                  'S/AHOKWA', 'ANKWASO', 'ANKWAASO', 'YAYASO', 'WASSA AGONA', 'SUHYENSO', 'S/ANHWIAM', 'ANHWIAM', 'WASSA ASIKUMA',
                  'HIAWA', 'HOTOPO', 'AWUDUA', 'TARKWA BONSO', 'AHEBENSO', 'KEJABRIL', 'AGONA AMENFI', 'WASA ABRESHIA',
                  'WASSA JAPA', 'SAYERANO', 'WASSA', 'KUBI', 'MFUOM', 'TARKWA  BOGOSO', 'AWASO', 'BAWDIE', 'ADJOAFUA', 'bogoso',
                  'SAMREBOI', 'ASANKRANGWA', 'ABASO', 'S/DWINASE', 'S/WIAWSO', 'WASSA AKYEMPIM', 'DWINASE', 'NAKABA',
                  'DOMEABRAH', 'DEBISO', 'KEDADWEN', 'BONSU', 'AMPAIN', 'AHWIAA', 'ADJAKAA MANSO', 'ASAWINSO',
                  'YAWMATWA', 'WASSA AKEMPEM', 'WASSA DADIESO', 'WASSA AKROPONG', 'NSAWORA', 'WASA KWAMANG', 'AXIM', 'PRESTEA', 'ASANKRAGUA',
                  'TAIN-NSAWKAW',]

NORTHERN_REGION = ['Tamale', 'Gusheigu', 'GUSHEGU', 'Bimbilla', 'Sagnerigu', 'Savelugu', 'Yendi', 'Karaga', 'Kpandai', 'Kumbungu',
                   'Sang', 'Nanton', 'Wulensi', 'Saboba', 'Tatale', 'Tolon', 'Zabzugu', 'Nalerigu', 'Gambaga', 'Walewale',
                   'Bunkpurugu', 'Chereponi', 'Yagaba', 'Yunyoo', 'Bole', 'Salaga', 'Damango', 'Sawla', 'Buipe', 'Daboya',
                   'Kpalbe', 'RICE CITY', 'TISHIGU', 'DALON', 'GUMBUGU', 'NYOLI', 'BUI', 'SAGNARIGU', 'CHAMBA', 'YAPEI',
                   'BANDA NKWANTA', 'JAMA', 'MPAHA', 'GUSHIEGU', 'SAWLA TUNA KALBA', 'NYANKPALA', 'BAMBOI', 'LARIBANGA',
                   'SAVELEGU', 'FUFULSO', 'MAKANGO', 'DAMONGO', 'ZABZEGU']

UPPER_EAST = ['Bawku', 'Bolga', 'Bolgatanga', 'Navrongo', 'Zebilla', 'Binduri', 'Zuarungu', 'Bongo', 'Sandema',
              'Fumbisi', 'Garu', 'Paga', 'Nangodi', 'Pusiga', 'Tongo', 'Tempane', 'BOLGAANGA', 'YIKENE',
              'PELUNGU', 'YIKENE  SOUTH', 'SIRIGU', 'NYARIGA', 'BUGRI', 'BAAZUA', 'WIDANA', 'BAZUA', 'kongo']

UPPER_WEST = ['Wa', 'Wa East', 'Wa West', 'Sissala East', 'Sissala', 'Sissala West', 'Lambussie', 'Jirapa', 'Nandom', 'Lawra', 'WA SOMBO', 'WECHIAU', 'POYENTANGA', 'ULLO', 'KABANYE', 'GWOLLU', 'LASSIA TUOLU', 'BIRIFOH',
              'kulpawn', 'KALEO', 'HAMILE', 'BABILE', 'TUMU', 'NADOWLI', 'NANDDOM', 'ESIAMA']

ACCRA_REGION = ['KORLE-BU', 'ACCRA', 'LEGON', 'CANTOMENTS', 'CANTONMENT', 'CANTOMENT', 'SPINTEX', 'DODOWA', 'TEMA', 'TAIFA', 'ACHIMOTA', 'ASHAIMAN', 'KASOA', 'TESHIE', 'NUNGUA',
                'DANSOMAN', 'LAPAZ', 'MADINA', 'OFANKOR', 'ABLEKUMA', 'AMASAMAN', 'AbekaLapaz', 'Abelemkpe', 'Ablekuma', 'Abokobi', 'Abossey Okai', 'Accra', 'Accra New Town', 'Achimota', 'Ada Foah',
                'Ada Kasseh', 'Adabraka', 'Adenta', 'Adenta East', 'Adjei Kojo', 'Adjen Kotoku', 'Airport City Accra', 'Airport Residential Area',
                'Akweteyman', 'Alajo', 'Amasaman', 'Apenkwa', 'Ashaiman', 'Ashongman', 'Asoprochona', 'Awoshie', 'Ayi Mensa', 'Bansa', 'Big Ada', 'Bubuashi',
                'Bubuashie', 'Bukom', 'Cantonments', 'Caprice', 'Chorkor', 'Christian Village', 'Dansoman', 'Darkuman', 'Dawhenya', 'Dodowa', 'Dome',
                'Dzorwulu', 'East Legon', 'Osu', 'Eleme', 'Gbawe', 'Gbegbe', 'Haatso', 'Kaneshie', 'Kisseman', 'Kokomlemle', 'Kokrobite', 'BURMA CAMP', 'ghana', 'Korle Gonno', 'Kotobabi', 'Kuntunse', 'Kwabenya', 'Kwashieman', 'Labadi', 'Labone, Accra', 'Lapaz', 'Lartebiokorshie', 'Lashibi', 'Lebanon Ashaiman',
                'Legon', 'Maamobi', 'Madina', 'Makola', 'Mallam', 'Mamprobi', 'Mataheko', 'Miotso', 'Nii Boi Town', 'Nima Accra', 'Nmai Dzorn', 'North Industrial Area', 'Nungua', 'Odorkor', 'Ogbodjo', 'Old Ningo', 'Oyarifa', 'Pantang', 'Pig Farm',
                'Pokuase', 'TRADE FAIR AR.', 'AIRPORT', 'ABEKA', 'ADJIRIGANO EAST LEGON', 'ASHALEY-BOTWE', 'Prampram', 'Sabon Zongo', 'Sakaman', 'Sakumono', 'Santeo', 'Sege', 'Shai Hills', 'Shiashie', 'Sowutuom', 'Sugbaniate', 'Swalaba', 'Taifa', 'Tema', 'Tema Community 4', 'Tema Community 5', 'Tema Manhean', 'Tesano', 'Teshie', 'Teshie  Nungua', 'Torkuase', 'Tudu', 'Weija', 'Korle Bu', 'kanda', 'kaneshi', 'Nima'
                ]

EASTERN_REGION = ['AKIM AKROSO', 'NKURAKAN', 'AKIM AWISA', 'AKIM-ODA', 'Abetifi', 'LARTEH-AKUAPIM', 'Abiriw', 'Abomosu', 'Aburi', 'Achiase', 'Adeiso', 'Adiemmra', 'Adoagyiri', 'Agormanya', 'Ahwerase', 'Akim Begoro', 'Akim Oda', 'Akim Swedru', 'Akim Tafo', 'Akosombo', 'AKIM APERADE', 'Akropong', 'Akuapem  Akropong', 'Akuapim  Mampong', 'Akuse', 'Akwamufie', 'Akwatia', 'Akyem  Awenare', 'Amanokrom', 'Anum', 'Anum  Boso', 'Anyinam', 'Apedwa', 'Aperadi', 'Asafo  Akyem', 'Asamankese', 'Asesewa', 'Asiakwa', 'Asona Town', 'Asuboe', 'Asuboni No.3', 'Asuboni Rails', 'Atibie, ', 'Atimpoku', 'Kade', 'Koforidua', 'Kibi', 'Kitase Akuapem', 'Koforidua  Asokore', 'Koforidua  Effiduase', 'Kpong', 'Krobo Odumase',
                  'Kukurantumi', 'AKIM ABOABO', 'Kwabeng', 'Kwahu Asafo', 'Kwahu Nsaba', 'Awukugua', 'Begoro', 'Bepong', 'Berekuso', 'Boso, ', 'Bososo', 'Coaltar', 'Dago, ', 'Dawu', 'Densuano', 'Donkorkrom', 'Ekowso', 'Enyiresi', 'Fodoa', 'Gyakiti', 'Hebron, ', 'Jejeti', 'Larteh Akuapem', 'Manhyia', 'Mpraeso', 'Mpraeso Amanfrom', 'New Abirem', 'Nkawkaw', 'Nkwatia Kwahu', 'Nsawam', 'Nuaso', 'Obo Kwahu', 'Oboo, ', 'Obosomase Akuapem', 'Old Tafo', 'Oseim', 'Osino', 'Otumi', 'Pakro', 'Pampanso', 'Peduase', 'Pepease', 'Pokrom', 'Senchi', 'Somanya', 'Suhum', 'Topremang', 'BUNSO'
                  ]
VOLTA_REGION = ['BATTOR', 'WORAWORA', 'Abor', 'Abutia Kpota', 'Abutia  Teti', 'Adafienu', 'Adaklu', 'Adaklu Waya', 'Adidome', 'Aflao', 'Agbozume', 'Agortime  Kpetoe',
                'Akatsi', 'Akome', 'Akpafu', 'Akrofu', 'Alakple', 'Alavanyo', 'Amedzofe, ', 'Anfoega', 'Anlo Afiadenyigba', 'Anloga', 'Anyako', 'Anyanui', 'Asukawkaw', 'Atiavi', 'Atimpoko', 'Ave  Dakpa', 'Aveyime  Battor', 'Baglo', 'Bame', 'Brewaniase', 'Dambai', 'Tefle', 'Dabala', 'Dafor', 'Denu', 'Dzelukope', 'Dzodze', 'Dzolokpuita', 'Gbefi', 'Gbledi  Agbogame', 'Hatsukope', 'Have', 'Hedzranawo', 'Hlefi', 'Ho', 'Hohoe', 'Juapong', 'Keta', 'Klefe', 'Klikor',
                'Kpale Kpalime', 'Kpalime Duga', 'Kpando', 'Kpedze', 'Kpeme', 'Kpetoe', 'Kpeve', 'Kpeve New Town', 'Leklebi', 'Logba Adzekoe', 'Lolobi', 'Mafi  Kumasi', 'Mepe', 'Nogokpo', 'Peki', 'Podoe', 'Seva, ', 'Shia, ', 'Sogakope', 'Sokpoe', 'Tadzewu', 'Tanyigbe', 'Taviefe', 'Tegbi', 'To Kpalime', 'Tokor', 'Tongor Kaira', 'Tsito', 'Vakpo', 'Vane, Avatime', 'Ve Golokwati', 'Ve  Koloenu', 'Vume', 'Wegbe Kpalime', 'Weta', 'Woe']

CENTRAL_REGION = ['CAPE-COAST', 'Abakrampa', 'Abandze', 'Abeadzi Kyiakor', 'Abrafo', 'Abura  Dunkwa', 'Adukrom', 'Afransi', 'Agona Abirem', 'Agona Swedru', 'Ajumako', 'Bamahu', 'swedru',
                  'Amissano', 'Ankaful', 'Anomabu', 'Anyimon', 'Apam', 'Asebu', 'Assin Asempaneye', 'Assin Darmang', 'Assin Fosu', 'Assin Kyekyewere', 'Assin Manso', 'Assin Foso', 'FOSO',
                  'Assin Nsuta', 'Assin Praso', 'Asutsuare', 'Asylum Down', 'Awutu Breku', 'Ayanfuri', 'Besease', 'Biriwa', 'Breman Asikuma', 'Buduburam', 'Cape Coast', 'Dawurampon', 'Diaso', 'Dunkwa  on  Offin', 'Effutu', 'Elmina', 'Esakyir', 'Gomoa Amoanda', 'Gomoa Mpota', 'Hemang', 'Jukwa', 'Jukwaa', 'Kasoa', 'Kwamankese', 'Kwanyako', 'Mankessim', 'Moree', 'Mozano', 'Mumford', 'Nsaba', 'Nyakrom', 'Nyankumase Ahenkro', 'Obrakyere', 'Odoben', 'Okyereko', 'Otuam', 'Potsin', 'Saltpond', 'Senya', 'Senya Beraku', 'Twifo Praso', 'Winneba', 'Yamoransa',
                  ]
NORTHERN_ZONE = {
    'KUMASI': KUMASI,
    'ASHANTI_REGION': ASHANTI_REGION,
    'WESTERN_REGION': WESTERN_REGION,
    'BONO_AHAFO_REGION': BONO_REGION,
    'NORTHERN_REGION': NORTHERN_REGION,
    'UPPER_EAST_REGION': UPPER_EAST,
    'UPPER_WEST_REGION': UPPER_WEST,
    'ACCRA_REGION': ACCRA_REGION,
    'EASTERN_REGION': EASTERN_REGION,
    'VOLTA_REGION': VOLTA_REGION,
    'CENTRAL_REGION': CENTRAL_REGION
}

idxs = []

copy_employed_df = employed_df
for region, towns in NORTHERN_ZONE.items():
    for town in towns:
        idx = employed_df.index[employed_df['FTOWN'].str.contains(
            town.upper(), na=False, regex=True)].tolist()
        idxs.extend(idx)
        employed_df = employed_df.drop(idx)

    df = copy_employed_df.iloc[idxs]
    idxs = []

    try:
        df.to_excel(f'{region}.xlsx')
    except Exception:
        print(f'Error: {Exception}')
    else:
        print(f'{region} Files Saved')

employed_df.to_excel('Updated.xlsx')
print('DONE!!!!')
