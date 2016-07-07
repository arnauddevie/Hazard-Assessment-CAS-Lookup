# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 17:17:45 2016
@author: Arnaud Devie
"""

#%% Data mining from Sigma Aldrich website
# Search URL by CAS number:
# http://www.sigmaaldrich.com/catalog/search?interface=CAS%20No.&term=1314-62-1&N=0&lang=en&region=US&focus=product&mode=mode+matchall
# On product page, Safety Information table, with H-statements, P-statements and PPE type

#==============================================================================
# Libraries
#==============================================================================
import re
import os
import sys
import time
import pandas
import urllib
from bs4 import BeautifulSoup
from selenium import webdriver

#==============================================================================
# Functions
#==============================================================================
def deblank(text):
    # Remove leading and trailing empty spaces
    return text.rstrip().lstrip()

def fixencoding(text):
    # Make string compatible with cp437 characters set (Windows console)
    return text.encode(encoding="cp437", errors="ignore").decode(encoding="utf-8", errors="ignore")

def deblankandcap(text):
    # Remove leading and trailing empty spaces, capitalize
    return text.rstrip().lstrip().capitalize()

def striphtml(text):
    # remove HTML tags from string (from: http://stackoverflow.com/a/3398894, John Howard)
    p = re.compile(r'<.*?>')
    return p.sub('', text)

def clean(text):
    # Deblank, fix encoding and strip HTML tags at once
    return striphtml(fixencoding(deblank(text)))

#==============================================================================
# Input
#==============================================================================
# Looking for info about chemical identified by CAS number ...
CASlist = list()
textfile = open('CAS-list.txt','r')
for line in textfile:
    CASlist.append(deblank(line.replace('\n','')))

textfile.close()

# Drop duplicates
CASlist = set(CASlist)

# Clean up
if '' in CASlist:
    CASlist.remove('')

# Override
#CASlist=['1120-71-4']

#%%
#==============================================================================
# Search patterns
#==============================================================================
Ppattern = '(P[0-9]{3}[0-9P\+]*)' # the letter P followed by 3 digits, including '+' combo
#Hpattern = 'H[0-9]{3}' # the letter H followed by 3 digits
Hpattern = '(H[0-9]{3}(?i)[ifd0-9H\+]*)' # the letter H followed by 3 digits, including '+' combo, case insensitive fd

# Parse H2P text file
textfile = open('H2P.txt', 'r')

# Initialize dictionary
H2P = dict()

for line in textfile:
    line = line.replace('\n','').replace('+ ','+') #.replace(',','')
    if re.match(Hpattern, line):
        hcode = re.match(Hpattern, line).group()
        H2P[hcode] = set(re.findall(Ppattern, line))

# Close textfile
textfile.close()

# Parse P-statements text file
textfile = open('P-statements.txt', 'r')

# Initialize dictionary
Pstatements = dict()

for line in textfile:
    line = line.replace('\n','').replace(' + ','+')
    if re.match(Ppattern, line):
        pcode = deblank(re.match(Ppattern, line).group())
        Pstatements[pcode] = deblank(line.split(pcode)[-1])

# Close textfile
textfile.close()

# Parse H-statements text file
textfile = open('H-statements.txt', 'r')

# Initialize dictionary
Hstatements = dict()

for line in textfile:
    line = line.replace('\n','').replace(' + ','+')
    if re.match(Hpattern, line):
        hcode = deblank(re.match(Hpattern, line).group())
        Hstatements[hcode] = deblank(line.split(hcode)[-1])

# Close textfile
textfile.close()

#==============================================================================
# Prevention, Response, Storage and Disposal P-statement from H-code
#==============================================================================
H2Prevention = dict()
H2Response = dict()
H2Storage = dict()
H2Disposal = dict()

for hcode in H2P:
    alist = H2Prevention.get(hcode,[])
    for pcode in H2P[hcode]:
        statement = Pstatements[pcode]
        if (pcode[1]=='2'): H2Prevention[hcode] = H2Prevention.get(hcode,[]); H2Prevention[hcode].append(statement)
        if (pcode[1]=='3'): H2Response[hcode]   = H2Response.get(hcode,[]); H2Response[hcode].append(statement)
        if (pcode[1]=='4'): H2Storage[hcode]    = H2Storage.get(hcode,[]); H2Storage[hcode].append(statement)
        if (pcode[1]=='5'): H2Disposal[hcode]   = H2Disposal.get(hcode,[]); H2Disposal[hcode].append(statement)

#%%
#==============================================================================
# Data mining Sigma Aldrich website
#==============================================================================

# Start Chrome instance
chromeOptions = webdriver.ChromeOptions()

if "SDS" not in os.listdir():
    os.mkdir("SDS")

prefs = {"download.default_directory" : os.path.join(os.getcwd(),"SDS"),
         "download.prompt_for_download" : False,
         "download.directory_upgrade" : True,
         "plugins.plugins_disabled" : ["Chrome PDF Viewer"]}
chromeOptions.add_experimental_option("prefs",prefs)
chromeOptions.add_argument("--disable-extensions")

if 'win' in sys.platform: # Windows
    chromedriver = os.path.join(os.getcwd(),'chromedriver','win32','chromedriver.exe')
elif 'darwin' in sys.platform: # Mac OS
    chromedriver = os.path.join(os.getcwd(),'chromedriver','mac32','chromedriver')
elif 'linux' in sys.platform: # Linux
    if sys.maxsize > 2**32: # 64-bit
        chromedriver = os.path.join(os.getcwd(),'chromedriver','linux64','chromedriver')
    else: # 32-bit
        chromedriver = os.path.join(os.getcwd(),'chromedriver','linux32','chromedriver')

driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)
driver.set_window_position(-2000, 0)

# Initialize
chemicals=list()
CASdict = dict()
badCAS = list()

for CAS in CASlist:

    chemical = dict()
    URL = dict()
    Name = ''

    # Store CAS #
    chemical['CAS'] = CAS
    print(CAS)

    try:
        # Webscraping search page
        searchURL = r'http://www.sigmaaldrich.com/catalog/search?interface=CAS%20No.&term=[INSERT-HERE]&N=0&lang=en&region=US&focus=product&mode=mode+matchall'.replace('[INSERT-HERE]',CAS)
        webpage = urllib.request.urlopen(searchURL).read()
        soup = BeautifulSoup(webpage, "html.parser")
        product = soup.find("li", class_='productNumberValue')
        productSubURL = product.a.decode().split('"')[1]
        sds = soup.find("li", class_='msdsValue')
        pattern = '\'(\w*)\'' # any string between ''
        [country, language, productNumber, brand] = re.findall(pattern, sds.a.get('href'))
        properties = soup.find("ul", class_="nonSynonymProperties")
        formula = striphtml(properties.span.decode_contents())

        # Webscraping product page
        productURL = 'http://www.sigmaaldrich.com[INSERT-HERE]'.replace('[INSERT-HERE]', productSubURL)
        webpage2 = urllib.request.urlopen(productURL).read()
        soup2 = BeautifulSoup(webpage2, "html.parser")

        # Store URLs
        chemical['SearchURL'] = searchURL
        chemical['ProductURL'] = productURL
        chemical['ProductNumber'] = productNumber
        chemical['Brand'] = brand
        chemical['Formula'] = formula


        # Name (compatible with cp437 characters set)
        Name = clean(soup2.find("h1", itemprop="name").decode_contents().split('\n')[1])
        chemical['Name'] = Name
        CASdict[CAS] = Name
        print(Name)

        # Synonyms
        try:
            Synonyms = [clean(synonym) for synonym in soup2.find("p", class_="synonym").findNext("strong").decode_contents().replace('\t','').replace('\n','').split(',')]
            chemical['Synonyms'] = Synonyms
        except:
            print('No Synonyms listed for %s - %s' % (CAS, Name))

        # List of H-statements
        soloHpattern = '(H[0-9]{3}(?i)[ifd]*)'
        try:
            codes = re.findall(soloHpattern, soup2.find("div", class_="safetyRight", id="Hazard statements").findNext("a", class_="ALL").decode_contents())
            statements = [Hstatements[code] for code in codes]
            Hazards = dict(zip(codes, statements))
            chemical['Hazards'] =  Hazards
        except:
            print('No Hazards listed for %s - %s' % (CAS, Name))

        # List of P-statements
        soloPpattern = '(P[0-9]{3})'
        try:
            codes = re.findall(Ppattern, soup2.find("div", class_="safetyRight", id="Precautionary statements").findNext("a", class_="ALL").decode_contents().replace(' ',''))
            statements = [' '.join([Pstatements[solo] for solo in re.findall(soloPpattern,code)]) for code in codes]
            Precautions = dict(zip(codes, statements))
            chemical['Precautions'] =  Precautions
        except:
            print('No Precautions listed for %s - %s' % (CAS, Name))

        # List of supplemental (non-GHS) H-statements
        try:
            suppstatements = soup2.find("div", class_="safetyRight", id="Supplemental Hazard Statements").decode_contents().split(',')
            chemical['Supp. Hazards'] =  [deblank(s) for s in set(suppstatements) if deblank(s) is not '']
        except:
            print('No supp. Hazards listed for %s - %s' % (CAS, Name))

        # List of PPE
        try:
            PPElist = soup2.find("div", class_="safetyRight", id="Personal Protective Equipment").findAll("a", class_="ALL")
            PPE = [deblank(ppe.decode_contents())[0].upper() + deblank(ppe.decode_contents())[1:] for ppe in PPElist]
            chemical['PPE'] = PPE
        except:
            print('No PPE listed for %s - %s' % (CAS, Name))

        # Download SDS as PDF file
        driver.get("http://www.sigmaaldrich.com/MSDS/MSDS/DisplayMSDSPage.do?country=%s&language=en&productNumber=%s&brand=%s" %(country, productNumber, brand));
        print("Downloading SDS file", end='')

        timedout = False
        timeout = time.time()
        while ("PrintMSDSAction.pdf" not in os.listdir('SDS')) and not timedout:
            print(".", end='')
            timeout = time.time() - timeout
            timedout = (timeout>30)
            time.sleep(1)

        if timedout:
            print(" Timed Out! Could not get the file")
        else:
            print(" Done.")
            sdsURL = os.path.join("SDS", Name + " - SDS.pdf")
            chemical['SDSfile'] = sdsURL
            os.rename(os.path.join("SDS","PrintMSDSAction.pdf"), sdsURL)

        # Store chemical
        chemicals.append(chemical)

    except:
        badCAS.append(CAS)
        print('Could not process %s - %s' % (CAS, Name))
        e = sys.exc_info()[0]

# Close Chrome instance
driver.quit()

# Display
print('Processed %d chemicals out of %d CAS numbers received' % (len(chemicals),len(CASlist)))

if len(badCAS) > 0:
    print('Unable to process the following CAS numbers:')
    for cas in badCAS: print(cas)

#%% Post processing
#==============================================================================
# Compilation of Statements
#==============================================================================
# Inventory of H-, P- and PPE statements
Hlist = list()
HfromCAS = dict()
HfromChemical = dict()

Plist = list()
PfromCAS = dict()
PfromChemical = dict()

PPElist=list()
PPEfromCAS = dict()
PPEfromChemical = dict()

Hsupplist = list()
HsuppfromCAS = dict()
HsuppfromChemical = dict()

for chemical in chemicals:
    if 'Hazards' in chemical.keys():
        for hazard in chemical['Hazards']:
            Hlist.append(hazard)
            alist = HfromCAS.get(hazard,[])
            alist.append(chemical['CAS'])
            alist = [item for item in set(alist)]
            alist.sort()
            HfromCAS[hazard] = alist

            alist = HfromChemical.get(hazard,[])
            alist.append(chemical['Name'])
            alist = [item for item in set(alist)]
            alist.sort()
            HfromChemical[hazard] = alist

    if 'Precautions' in chemical.keys():
        for precaution in chemical['Precautions']:
            Plist.append(precaution)
            alist = PfromCAS.get(precaution,[])
            alist.append(chemical['CAS'])
            alist = [item for item in set(alist)]
            alist.sort()
            PfromCAS[precaution] = alist

            alist = PfromChemical.get(precaution,[])
            alist.append(chemical['Name'])
            alist = [item for item in set(alist)]
            alist.sort()
            PfromChemical[precaution] = alist

    if 'PPE' in chemical.keys():
        for ppe in chemical['PPE']:
            PPElist.append(ppe)
            alist = PPEfromCAS.get(ppe,[])
            alist.append(chemical['CAS'])
            alist = [item for item in set(alist)]
            alist.sort()
            PPEfromCAS[ppe] = alist

            alist = PPEfromChemical.get(ppe,[])
            alist.append(chemical['Name'])
            alist = [item for item in set(alist)]
            alist.sort()
            PPEfromChemical[ppe] = alist

    if 'Supp. Hazards' in chemical.keys():
        for hazard in chemical['Supp. Hazards']:
            Hsupplist.append(hazard)
            alist = HsuppfromCAS.get(hazard,[])
            alist.append(chemical['CAS'])
            alist = [item for item in set(alist)]
            alist.sort()
            HsuppfromCAS[hazard] = alist

            alist = HsuppfromChemical.get(hazard,[])
            alist.append(chemical['Name'])
            alist = [item for item in set(alist)]
            alist.sort()
            HsuppfromChemical[hazard] = alist

# Count instances of each H-statement
Hdict = dict()
for Hstatement in Hlist:
    key = Hstatement
    Hdict[key] = Hdict.get(key, 0) + 1

# Count instances of each P-statement
Pdict = dict()
for Pstatement in Plist:
    key = Pstatement
    Pdict[key] = Pdict.get(key, 0) + 1

# Count instances of each PPE recommendation
PPEdict = dict()
for ppe in PPElist:
    key=ppe
    PPEdict[key] = PPEdict.get(key, 0) + 1

# Count instances of each supplemental Hazard statement
Hsuppdict = dict()
for statement in Hsupplist:
    key = statement
    Hsuppdict[key] = Hsuppdict.get(key, 0) + 1

# Create a dataframe with a list of unique H-statements
H = pandas.DataFrame(Hlist, columns = {'Code'})
Hunique = H[H.Code!=''].drop_duplicates()

Hunique['Count']            = Hunique['Code'].map(Hdict)
Hunique['Statement']        = Hunique['Code'].map(Hstatements)
Hunique['Assoc.Pcode']      = Hunique['Code'].str.slice(0,4).map(H2P)
Hunique['Assoc.CAS']        = Hunique['Code'].map(HfromCAS)
Hunique['Assoc.Chemical']   = Hunique['Code'].map(HfromChemical)
Hunique['Prevention']       = Hunique['Code'].str.slice(0,4).map(H2Prevention)
Hunique['Response']         = Hunique['Code'].str.slice(0,4).map(H2Response)
Hunique['Storage']          = Hunique['Code'].str.slice(0,4).map(H2Storage)
Hunique['Disposal']         = Hunique['Code'].str.slice(0,4).map(H2Disposal)

# Create a dataframe with a list of unique P-statements
P = pandas.DataFrame(Plist, columns = {'Code'})
Punique = P[P.Code!=''].drop_duplicates()

codes = Punique['Code']
for code in codes:
    statements = [' '.join([Pstatements[solo] for solo in re.findall(soloPpattern,code)]) for code in Punique['Code']]
Precautions = dict(zip(codes, statements))

Punique['Count']            = Punique['Code'].map(Pdict)
Punique['Statement']        = Punique['Code'].map(Precautions)
Punique['Assoc.CAS']        = Punique['Code'].map(PfromCAS)
Punique['Assoc.Chemical']   = Punique['Code'].map(PfromChemical)

# Create a dataframe with a list of unique PPE requirements
PPE = pandas.DataFrame(PPElist, columns = {'Item'})
PPEunique = PPE[PPE.Item!=''].drop_duplicates()

PPEunique['Count']            = PPEunique['Item'].map(PPEdict)
PPEunique['Assoc.CAS']        = PPEunique['Item'].map(PPEfromCAS)
PPEunique['Assoc.Chemical']   = PPEunique['Item'].map(PPEfromChemical)

# Create a dataframe with a list of unique supplemental hazards
Hsupp = pandas.DataFrame(Hsupplist, columns = {'Statement'})
Hsuppunique = Hsupp[Hsupp.Statement!=''].drop_duplicates()

codes = list()
statements = Hsuppunique['Statement']
for idx, statement in enumerate(statements):
    codes.append('Supp. %d' % (idx+1))
Hsuppcodes = dict(zip(statements, codes))

Hsuppunique['Code']            = Hsuppunique['Statement'].map(Hsuppcodes)
Hsuppunique['Count']            = Hsuppunique['Statement'].map(Hsuppdict)
Hsuppunique['Assoc.CAS']        = Hsuppunique['Statement'].map(HsuppfromCAS)
Hsuppunique['Assoc.Chemical']   = Hsuppunique['Statement'].map(HsuppfromChemical)

# Concatenate GHS Hazards and supplemental Hazards
Hcombo = pandas.concat([Hunique, Hsuppunique])

#==============================================================================
# Table of all chemicals
#==============================================================================
chemicalsDF = pandas.DataFrame(chemicals)
chemicalsDF['Product Number'] = chemicalsDF.apply(lambda row: '<a href="' + row['ProductURL'] + '">' + row['ProductNumber'] + '</a>',axis=1)
chemicalsDF['SDS'] = chemicalsDF.apply(lambda row: '<a href="' + row['SDSfile'] + '">SDS</a>',axis=1)


#%% Export
# HTML table settings
pandas.set_option('display.max_colwidth', -1)

# List of chemicals, sorted by name, with CAS, URL, SDS,...
inventory = open("Inventory.html",'w')
inventory.write(chemicalsDF.sort_values('Name').to_html(index=False, na_rep='-', escape=False, columns=['Name', 'Synonyms', 'CAS', 'Formula', 'Hazards', 'Precautions', 'PPE', 'Product Number', 'SDS']))
inventory.close()

# Export Hazards and supplemental Hazards to HTML file
Hlist = open('Hlist.html','w')
Hlist.write(Hcombo.sort_values('Code').to_html(index=False, na_rep='-', columns=['Code', 'Count', 'Statement', 'Assoc.Chemical', 'Prevention', 'Response', 'Storage', 'Disposal']))
Hlist.close()

# Export Precautions to HTML file
Plist = open('Plist.html','w')
Plist.write(Punique.sort_values('Code').to_html(index=False, na_rep='-', columns=['Code', 'Count', 'Statement', 'Assoc.Chemical']))
Plist.close()

# Export PPE to HTML file
PPElist = open('PPElist.html','w')
PPElist.write(PPEunique.sort_values('Item').to_html(index=False, na_rep='-', columns=['Item', 'Count', 'Assoc.Chemical']))
PPElist.close()

## Export supplemental Hazards to HTML file
#Slist = open('Slist.html','w')
#Slist.write(Hsuppunique.sort_values('Code').to_html(index=False, columns=['Code', 'Count', 'Statement', 'Assoc.Chemical'], na_rep='-'))
#Slist.close()

# Export to Excel file (via xlsxwriter)
writer = pandas.ExcelWriter('Hazard Assessment.xlsx', engine='xlsxwriter')
Hcombo.sort_values('Code').to_excel(writer,'Hazard Assessment', index=False, na_rep='-')
Punique.sort_values('Code').to_excel(writer,'Precautions', index=False, na_rep='-')
PPEunique.sort_values('Item').to_excel(writer,'PPE', index=False, na_rep='-')
writer.save()
