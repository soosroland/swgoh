from csv import writer
import xlsxwriter
import pandas as pd
import numpy as np

# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('FejlesztesiTerv.xlsx')
 
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()

unit_current = ['TRIPLEZERO','50RT', 'AAYLASECURA','ADMIRALACKBAR','ADMIRALPIETT','ADMIRALRADDUS','AHSOKATANO','FULCRUMAHSOKA',
'AMILYNHOLDO','ARCTROOPER501ST','ASAJVENTRESS','AURRA_SING','B1BATTLEDROIDV2','B2SUPERBATTLEDROID','BARRISSOFFEE','BASTILASHAN',
'BASTILASHANDARK','BAZEMALBUS','BB8','BENSOLO','BIGGSDARKLIGHTER','BISTAN','BOKATAN','BOBAFETT','BOBAFETTSCION','BODHIROOK','BOSSK',
'BOUSHH','BT1','C3POLEGENDARY','CADBANE','CANDEROUSORDO','HOTHHAN','PHASMA','CARADUNE','CARTHONASI','CASSIANANDOR','CC2224',
'CHEWBACCALEGENDARY','CHIEFCHIRPA','CHIEFNEBIT','CHIRRUTIMWE','CHOPPERS3','CLONESERGEANTPHASEI','CLONEWARSCHEWBACCA','COLONELSTARCK',
'COMMANDERAHSOKA','COMMANDERLUKESKYWALKER','CORUSCANTUNDERWORLDPOLICE','COUNTDOOKU','CT210408','CT5555','CT7567','DARKTROOPER',
'DARTHMALAK','DARTHMALGUS','MAUL','DARTHNIHILUS','DARTHREVAN','DARTHSIDIOUS','DARTHSION','DARTHTALON','DARTHTRAYA','VADER','DASHRENDAR',
'DATHCHA','DEATHTROOPER','DENGAR','DIRECTORKRENNIC','DROIDEKA','BADBATCHECHO','EETHKOTH','EIGHTHBROTHER','EMBO','EMPERORPALPATINE',
'ENFYSNEST','EWOKELDER','EWOKSCOUT','EZRABRIDGERS3','FENNECSHAND','FIFTHBROTHER','FINN','FIRSTORDEREXECUTIONER','FIRSTORDEROFFICERMALE',
'FIRSTORDERSPECIALFORCESPILOT','FIRSTORDERTROOPER','FIRSTORDERTIEPILOT','GAMORREANGUARD','GARSAXON','ZEBS3','GRIEVOUS','GENERALHUX','GENERALKENOBI',
'GENERALSKYWALKER','VEERS','GEONOSIANBROODALPHA','GEONOSIANSOLDIER','GEONOSIANSPY','GRANDADMIRALTHRAWN','GRANDINQUISITOR','GRANDMASTERYODA',
'GRANDMOFFTARKIN','GREEDO','GREEFKARGA','HANSOLO','HERASYNDULLAS3','HERMITYODA','HK47','HONDO','HOTHREBELSCOUT','HOTHREBELSOLDIER',
'BADBATCHHUNTER','IDENVERSIOEMPIRE','MAGNAGUARD','IG11','IG86SENTINELDROID','IG88','IMAGUNDI','IMPERIALPROBEDROID','IMPERIALSUPERCOMMANDO',
'JABBATHEHUTT','JANGOFETT','JAWA','JAWAENGINEER','JAWASCAVENGER','JEDIKNIGHTCONSULAR','ANAKINKNIGHT','JEDIKNIGHTGUARDIAN','JEDIKNIGHTLUKE',
'JEDIKNIGHTREVAN','JEDIMASTERKENOBI','GRANDMASTERLUKE','JOLEEBINDO','JUHANI','JYNERSO','K2SO','KANANJARRUSS3','KIADIMUNDI','KITFISTO','KRRSANTAN',
'KUIIL','KYLEKATARN','KYLOREN','KYLORENUNMASKED','L3_37','ADMINISTRATORLANDO','LOBOT','LOGRAY','LORDVADER','LUKESKYWALKER','LUMINARAUNDULI',
'MACEWINDU','MAGMATROOPER','MARAJADE','MAULS7','MISSIONVAO','HUMANTHUG','MOFFGIDEONS1','MONMOTHMA','MOTHERTALZIN','NIGHTSISTERACOLYTE',
'NIGHTSISTERINITIATE','NIGHTSISTERSPIRIT','NIGHTSISTERZOMBIE','NINTHSISTER','NUTEGUNRAY','OLDBENKENOBI','DAKA','BADBATCHOMEGA','PADMEAMIDALA',
'PAO','PAPLOO','PLOKOON','POE','POGGLETHELESSER','PRINCESSLEIA','QIRA','QUIGONJINN','R2D2_LEGENDARY','RANGETROOPER','HOTHLEIA','EPIXFINN','EPIXPOE',
'RESISTANCEPILOT','RESISTANCETROOPER','GLREY','REYJEDITRAINING','REY','ROSETICO','ROYALGUARD','SABINEWRENS3','SANASTARROS','SAVAGEOPRESS',
'SCARIFREBEL','SECONDSISTER','SEVENTHSISTER','SHAAKTI','SHORETROOPER','SITHASSASSIN','SITHTROOPER','SITHPALPATINE','SITHMARAUDER','FOSITHTROOPER',
'UNDERCOVERLANDO','SNOWTROOPER','STARKILLER','STORMTROOPER','STORMTROOPERHAN','SUNFAC','SUPREMELEADERKYLOREN','T3_M4','TALIA','BADBATCHTECH','TEEBO',
'ARMORER','THEMANDALORIAN','THEMANDALORIANBESKARARMOR','THIRDSISTER','C3POCHEWBACCA','TIEFIGHTERPILOT','TUSKENRAIDER','TUSKENSHAMAN','UGNAUGHT',
'URORRURRR','YOUNGCHEWBACCA','SMUGGLERCHEWBACCA','SMUGGLERHAN','VISASMARR','WAMPA','WATTAMBOR','WEDGEANTILLES','WICKET', 'BADBATCHWRECKER','YOUNGHAN',
'YOUNGLANDO','ZAALBAR','ZAMWESELL']

unit_req_p1 = [4,0, 0,0,1,0,0,1,
1,1,1,0,1,0,1,0,
0,0,6,2,0,0,0,1,5,0,1,
0,3,6,0,2,0,0,0,0,0,1,
5,0,0,0,0,1,0,1,
3,0,0,1,0,0,0,1,
0,2,3,0,10,2,3,0,10,3,0,
0,0,1,0,2,0,1,0,1,5,
0,0,0,0,0,1,0,0,2,
2,0,0,1,0,0,5,0,4,
10,0,2,0,0,4,0,4,
0,0,1,5,0,7,0,1,0,0,
1,0,0,0,1,2,1,2,1,
0,1,0,1,0,0,0,0,8,
4,0,0,2,0,0,0,0,3,0,0,
1,0,0,0,1,0,1,3,0,0,0,
0,0,1,0,0,0,0,0,0,0,
1,0,0,1,1,0,0,0,2,
0,1,0,1,0,0,0,0,4,0,1,0,0,
0,0,0,3,0,0,0,1,5,0,
0,0,0,0,0,1,0,0,0,0,
0,0,0,0,0,0,0,1,1,0,1,
1,0,7,0,3,1,0,1,3,
0,0,0,0,1,0,6,0,0,0,0,
0,1,0]

unit_req_p2 = [0,0,2,1,0,1,0,0,
0,1,1,0,0,2,0,1,
0,0,2,13,0,1,0,0,12,1,0,
1,1,2,1,0,0,0,0,0,0,0,
2,0,0,0,0,1,0,1,
11,1,0,1,1,1,0,0,
4,10,0,0,6,0,0,0,0,0,0,
0,2,1,0,1,0,0,0,2,4,
1,0,0,1,0,0,0,0,2,
0,0,0,0,0,0,1,1,1,
4,2,0,0,0,4,6,4,
0,0,2,1,0,1,2,0,0,0,
0,0,0,0,0,0,0,4,0,
0,0,1,0,0,0,0,0,3,
1,2,0,0,1,0,0,2,1,0,0,
1,0,0,2,2,1,0,1,1,0,0,
0,1,0,8,0,0,0,1,0,0,
0,0,0,0,0,0,0,1,1,
0,0,0,0,0,0,0,0,2,0,0,0,0,
0,0,3,2,0,1,0,0,1,0,
0,0,0,0,0,0,0,4,0,1,
0,1,4,0,0,0,3,0,0,0,1,
0,0,1,0,2,0,0,0,0,
0,1,1,0,0,0,4,1,0,0,1,
0,1,0]

unit_req_p3 = [1,1,0,0,0,0,0,0,
0,0,0,1,1,0,1,0,
2,0,3,7,0,0,0,0,10,1,0,
0,0,3,1,0,0,0,1,0,0,0,
0,0,0,1,0,1,0,0,
9,3,1,0,0,1,0,1,
2,14,1,0,1,0,0,0,1,0,0,
0,0,2,1,0,0,0,0,1,0,
0,1,1,0,0,1,0,0,0,
1,0,0,0,1,0,3,1,1,
1,0,0,0,1,1,0,2,
1,0,0,2,0,2,2,1,2,1,
1,0,0,0,0,1,0,2,0,
0,0,0,0,0,0,1,0,4,
2,5,7,0,0,0,1,0,3,0,0,
0,0,0,1,0,0,0,5,6,0,1,
1,0,1,11,0,0,0,0,1,1,
0,0,0,0,0,2,0,0,3,
0,0,0,0,0,0,0,1,4,0,2,1,1,
0,2,7,1,0,1,0,0,0,1,
0,0,1,0,0,1,0,7,0,1,
2,0,2,0,0,0,4,0,1,0,0,
0,0,5,0,0,0,2,0,0,
0,0,0,0,0,0,3,0,0,0,1,
0,0,0]

unit_req_p4 = [0,3, 1,1,0,0,0,0,
0,0,3,2,0,1,0,0,
1,1,1,6,0,2,1,0,6,0,0,
0,1,0,1,0,0,0,0,0,0,0,
2,0,1,0,0,0,0,0,
7,0,1,0,2,0,0,1,
6,4,1,0,0,0,0,0,0,1,1,
1,0,1,1,0,0,1,1,1,0,
0,0,0,0,0,0,1,0,0,
0,1,1,1,0,0,0,0,3,
4,2,0,1,0,0,0,1,
0,0,0,2,0,5,1,2,0,0,
0,1,0,0,0,2,0,0,0,
0,0,0,0,0,1,0,0,4,
2,3,8,1,0,0,0,0,2,0,1,
0,1,1,0,0,0,0,2,7,0,0,
0,0,1,6,0,0,0,0,0,3,
1,1,1,0,0,0,0,0,2,
0,0,0,0,1,0,0,0,1,1,4,0,0,
0,0,7,1,0,0,0,0,1,1,
2,1,2,0,0,0,1,9,0,0,
0,2,13,0,0,0,7,0,0,1,0,
0,0,2,0,0,1,0,0,0,
0,0,0,1,0,1,0,1,1,0,0,
0,2,0]

unit_req_p5 = [0,3, 0,0,0,0,0,1,
1,0,0,0,0,0,2,2,
0,2,0,2,0,1,0,1,9,2,0,
1,0,2,0,0,0,0,0,0,0,0,
2,0,0,1,1,1,0,1,
6,3,0,1,1,0,1,0,
9,7,0,0,0,1,0,0,0,0,0,
0,0,0,0,0,1,2,0,0,1,
0,0,0,0,0,0,0,1,1,
2,0,1,0,0,0,0,0,9,
10,0,0,1,0,0,10,5,
0,0,0,0,0,0,0,1,0,1,
1,1,1,0,0,3,1,0,0,
0,1,0,0,0,0,0,0,12,
3,0,8,1,0,0,0,0,0,0,0,
0,0,2,1,0,1,1,0,5,0,0,
1,1,0,4,0,0,1,1,0,0,
0,0,1,0,0,0,2,1,2,
0,0,0,0,0,0,0,0,2,0,0,1,0,
1,0,5,0,0,1,0,0,0,0,
0,0,0,0,0,1,2,8,0,0,
0,1,13,0,0,0,4,1,2,2,0,
1,0,0,0,0,0,0,2,0,
0,0,1,0,0,1,0,3,1,0,0,
0,1,0]

unit_req_p6 = [0,0,0,0,0,1,0,0,
0,0,2,1,2,1,0,6,
0,0,1,10,1,0,0,0,9,0,0,
0,0,0,0,0,0,0,0,0,0,0,
1,0,1,0,0,1,0,0,
9,0,1,2,0,0,0,1,
0,11,1,0,0,0,0,1,3,0,5,
0,0,1,0,0,0,0,0,0,0,
0,0,0,0,0,1,0,0,1,
0,0,0,0,0,0,0,1,5,
0,0,2,1,1,0,0,1,
1,0,0,6,0,5,2,3,0,0,
0,0,2,0,1,2,2,2,2,
0,0,1,1,1,1,0,0,0,
1,7,9,0,0,0,0,1,3,0,1,
0,0,0,1,1,0,1,0,10,0,0,
0,0,2,8,0,0,0,0,0,2,
1,1,0,2,1,0,1,0,1,
1,0,0,1,0,0,0,0,0,1,4,0,0,
1,0,11,0,2,0,1,0,2,1,
0,0,0,0,1,0,1,3,0,2,
2,0,0,0,0,0,8,0,0,1,0,
0,0,1,0,0,0,1,0,2,
0,0,1,0,0,0,2,0,0,0,0,
1,1,1]


file_name = "E:\\Roland\\swgoh\\csvs\\unit-export.csv"

# write header
worksheet.write(0, 0, "Character name")
worksheet.write(0, 1, "R5 current")
worksheet.write(0, 2, "P1 req")
worksheet.write(0, 3, "Needed number")
worksheet.write(0, 4, "Closest to req")


row_num = 1
for i in range(len(unit_current)):
    #18. Ben Solo 23. Boba Scion 54. Darth Malgus 119. Jabba 147. LV 153. Maul 184. Rey 190. Sana 199. SEE 204. SK
    if i != 19 and i != 24 and i != 55 and i != 120 and i != 148 and i != 154 and i != 185 and i != 191 and i != 200 and i != 205:
        # reading the csv file
        df = pd.read_csv(file_name, encoding="ISO-8859-1", on_bad_lines='skip', sep=',')

        # write if we do not have enough for P1
        if ( df[df['Relic Tier'] >= 5]['Relic Tier'].count() < unit_req_p1[i] ):

            column = 0
            # write character name
            worksheet.write(row_num, column, str(unit_current[i]))
            column = column + 1 

            # write R5 data
            worksheet.write(row_num, column, str(df[df['Relic Tier'] >= 5]['Relic Tier'].count()))
            column = column + 1 

            # write P1 required
            worksheet.write(row_num, column, unit_req_p1[i])
            column = column + 1 

            # how many more is needed from that unit
            more_needed = unit_req_p1[i] - df[df['Relic Tier'] >= 5]['Relic Tier'].count()
            # write P1 more needed
            worksheet.write(row_num, column, more_needed)
            column = column + 1
            # data with lower relic characters
            data_withlowerrelic = df[df['Relic Tier'] < 5]
            # data with lower relic characters sorted based on relic level
            data_withlowerrelic_sorted = data_withlowerrelic.sort_values(by=['Relic Tier'], ascending=False)
            # getting player names with relic levels
            players_withlowerrelic = data_withlowerrelic_sorted[['ï»¿Player Name', 'Relic Tier']]
            # only leave first X players (where is is needed character number + 2)
            players_withlowerrelic = players_withlowerrelic.iloc[0:more_needed+2]
            # print out player names to output
            message = ""
            for idx in range(len(players_withlowerrelic.index)):
                message += str(players_withlowerrelic['ï»¿Player Name'].iloc[idx]) + ": R" + str(int(players_withlowerrelic['Relic Tier'].iloc[idx])) + ", "
            worksheet.write(row_num, column, message)
            column = column + 1
            row_num = row_num + 1

    elif i == 19 or i == 24 or i == 55 or i == 120 or i == 148 or i == 154 or i == 185 or i == 191 or i == 200 or i == 205:
        # reading the csv file
        df = pd.read_csv(file_name, encoding="ISO-8859-1", on_bad_lines='skip', sep=',')

        # write if we do not have enough for P1
        if ( df[df['Gear Tier'] >= 5]['Gear Tier'].count() < unit_req_p1[i] ):

            column = 0
            # write character name
            worksheet.write(row_num, column, str(unit_current[i]))
            column = column + 1 

            # write R5 data
            worksheet.write(row_num, column, str(df[df['Gear Tier'] >= 5]['Gear Tier'].count()))
            column = column + 1 

            # write P1 required
            worksheet.write(row_num, column, unit_req_p1[i])
            column = column + 1 

            # how many more is needed from that unit
            more_needed = unit_req_p1[i] - df[df['Gear Tier'] >= 5]['Gear Tier'].count()
            # write P1 more needed
            worksheet.write(row_num, column, more_needed)
            column = column + 1
            # data with lower relic characters
            data_withlowerrelic = df[df['Gear Tier'] < 5]
            # data with lower relic characters sorted based on relic level
            data_withlowerrelic_sorted = data_withlowerrelic.sort_values(by=['Gear Tier'], ascending=False)
            # getting player names with relic levels
            players_withlowerrelic = data_withlowerrelic_sorted[['ï»¿Player Name', 'Gear Tier']]
            # only leave first X players (where is is needed character number + 2)
            players_withlowerrelic = players_withlowerrelic.iloc[0:more_needed+2]
            # print out player names to output
            message = ""
            for idx in range(len(players_withlowerrelic.index)):
                message += str(players_withlowerrelic.index.values[idx]) + ": R" + str(int(players_withlowerrelic['Gear Tier'].iloc[idx])) + ", "
            worksheet.write(row_num, column, message)
            column = column + 1
            row_num = row_num + 1

    file_name = "E:\\Roland\\swgoh\\csvs\\unit-export(" + str(i+1) + ").csv"
# close the file
workbook.close()