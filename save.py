from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time

browser = webdriver.Firefox()
browser.get("https://swgoh.gg/g/ly2CVgc8SlqNtmjLEJ07ng/unit-search/")

# Maximize the window and let code stall 
# for 10s to properly maximise the window.
browser.maximize_window()
time.sleep(5)

# Get unit names from dropdown
units = browser.find_elements(By.TAG_NAME, 'select')
unit_list = units[0].text
unit_split = unit_list.split("\n")
unit_split.pop(0) # remove "-- Select Unit --"
# Get dropdown element
selected = Select(units[0])
# not working unit name selector..
#selected.select_by_value(units[1].text)
#units = E:\Roland\swgoh\units.txt read in the unit names
# Loop through units
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
'URORRURRR','YOUNGCHEWBACCA','SMUGGLERCHEWBACCA','SMUGGLERHAN','VISASMARR','WAMPA','WATTAMBOR','WEDGEANTILLES','WICKET','BADBATCHWRECKER','YOUNGHAN',
'YOUNGLANDO','ZAALBAR','ZAMWESELL']
#unit_current = unit_split
for i in range(len(unit_current)):
    #select.select_by_value('TRIPLEZERO')
    selected.select_by_value(unit_current[i])
    unit = browser.find_elements(By.TAG_NAME, 'a')
    for e in unit:
        #print(e.text)
        if e.text == "CSV Export":
            browser.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            e.click()
            time.sleep(1)

#for e in units:
    #print(e.text)
#unit.click()
#elem = browser.find_elements_by_xpath("//*[@type='submit']")#put here the content you have put in Notepad, ie the XPath
#elem = browser.find_element("xpath", '/html/body/div[3]/div[1]/div[2]/ul/div/div/div/li/div/div[1]/div/div/div/div[2]/select')#put here the content you have put in Notepad, ie the XPath
#elem = browser.find_element("xpath", '/html/body/div[3]/div[1]/div[2]/ul/div/div/div/li[1]/div/div[1]/div/div/div/div/select')#.sendKeys("TRIPLEZERO")
#button = browser.find_element("id",'buttonID') #Or find button by ID.
#print(elem.get_attribute("class"))
unit = browser.find_elements(By.TAG_NAME, 'a')
for e in unit:
    #print(e.text)
    if e.text == "CSV Export":
        browser.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        e.click()
print("done")
time.sleep(30)
browser.close()