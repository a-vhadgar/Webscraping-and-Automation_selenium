from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import NoSuchElementException
import sys
import os
from bs4 import BeautifulSoup
import requests
import urllib
import pandas as pd
import subprocess
import openpyxl
import numpy as np
import datetime as dt
import obc as o
#getting path of current directory
path=os.path.dirname(os.path.abspath(__file__))


def getting_FOA(foa,chromedriver_path):
    """ retrives precceding FOA for given FOA """
    if foa.startswith('PAR'):
        link= 'https://grants.nih.gov/grants/guide/pa-files/'
        newlink=link+foa+'.html'
        new_list=[]
        new_list.append(foa)
        url1=requests.get(newlink)
        soup=BeautifulSoup(url1.text, 'html.parser')
        hit=soup.findAll('a')
        if hit[8].text=='':
            pass
        else:
            new_list.append(hit[8].text)
    else:
        link= 'https://grants.nih.gov/grants/guide/rfa-files/'
        newlink=link+foa+'.html'
        new_list=[]
        new_list.append(foa)
        url1=requests.get(newlink)
        soup=BeautifulSoup(url1.text, 'html.parser')
        hit=soup.findAll('a')
        if hit[3].text=='':
            pass
        else:
            new_list.append(hit[3].text)
    foa_number=','.join(new_list)
    return foa_number


def download_excel(foa_number):
    """Automatically fetch and downloads grants details
        for both FOA using Selenium  from https://projectreporter.nih.gov/reporter.cfm """
    link=""
    try:
        options = Options()
        prefs={"download.default_directory" : path}
        options.add_experimental_option("prefs",prefs)
        options.headless = False
        driver = webdriver.Chrome(options=options, executable_path= chromedriver_path)
        driver.get('https://projectreporter.nih.gov/reporter.cfm')
        driver.find_element_by_xpath('//*[@id="formDiv"]/div[3]/div[2]/div/div[3]/div/form/div[1]/div[2]/fieldset/legend/a').click()
        driver.find_elements_by_xpath('//*[@id="divOptionsFY"]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]')[0].click()
        driver.find_element_by_xpath('//*[@id="formDiv"]/div[3]/div[2]/div/div[3]/div/form/div[1]/div[2]/fieldset/legend/a').click()
        foa=driver.find_element_by_id('p_RFA')
        foa.send_keys(foa_number)
        driver.find_element_by_xpath('//*[@title="Submit Query"]').click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,"ExportDiv")))
        driver.switch_to.window(driver.window_handles[-1])
        driver.find_element_by_xpath('//*[@title="Go"]').click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.find_element_by_xpath('//*[@title="All"]').click()
        driver.find_elements_by_css_selector("input[type='radio'][value='csv']")[0].click()
        driver.find_elements_by_xpath('//*[@id="divOptions"]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]')[0].click()
        driver.find_elements_by_xpath('//*[@id="divOptions"]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/table/tbody/tr[3]')[0].click()
        driver.find_element_by_xpath('//*[@title="Export"]').click()
        driver.implicitly_wait(30)
        driver.switch_to.window(driver.window_handles[-1])
        iframe=driver.find_element_by_tag_name("iframe")
        driver.switch_to.frame(iframe)
        link=driver.find_element_by_xpath('//*[@id="processingmessage"]/div/table/tbody/tr/td/table/tbody/tr/td/a').get_attribute('href')
        val1=driver.find_element_by_xpath('//*[@id="processingmessage"]/div/table/tbody/tr/td/table/tbody/tr/td/a').click()
        driver.implicitly_wait(40)
    except NoSuchElementException:
        print("element not found, please re-run the script")
    finally:
        time.sleep(10)
        local_filename = link.split('/')[-1]
        file_path=path+'\\'+local_filename
        driver.quit()
    if os.path.isfile(file_path):
        print("file has been sucessfully downloaded at {0} ".format(file_path))
    else:
        print("re-run the script")
    return file_path


def auto_incrementNumber(number):
    """ increments grants index """
    number=number+1
    return number

def grants_list(file_path):
    """ discards grants other thatn type 1 and saves grants list to text file """ 
    today = dt.datetime.today().strftime("%Y-%m-%d")
    df=pd.read_csv(file_path,skiprows=4)
    df=df[df.Type==1]
    df['Grant identifier']=df['Project Number'].apply(lambda x: x[4:].split("-")[0])
    df['Project Number']=df['Project Number'].apply(lambda x: x[1:].split("-")[0])
    dict1={'NIGMS':'National Institute for General Medical Sciences','NIAID':'National Institute of Allergy and Infectious Diseases'}
    df['Funding Organization']=df['Funding IC'].map(dict1)
    dflist=df['Grant identifier'].tolist()
    list1=[ x.lower()+"[gr]" for x in dflist]
    filename=path+'\\'+'grants{}.txt'.format(today)
    df.to_csv('out11.csv',index=False)
    with open(filename, 'w') as f:
        for item in list1:
            f.write("%s\n" % item)
    return df,filename

def ttl_file(df,number1,dict1):
    """ Fetches pubmid for grants from XML file genereted from search_grant.py file and creates .ttl file for grants which are not in obc.ide."""
    list1=[]
    for l in df['Grant identifier']:
        for pubid in dict1:
                if l in pubid:
                    print(l)
                    list1.append([l,dict1[pubid]])
    df1 = pd.DataFrame(list1,columns = ['Grant identifier','pubid'])
    df2=df.merge(df1,on='Grant identifier')
    print(df2['pubid'][0])
    rdf=[]
    today = dt.datetime.today().strftime("%Y-%m-%d")
    email_id='akshata.vhadgar@ufl.edu'
    df2.to_csv('out.csv',index=False)
    
    for index, row in df2.iterrows():
        title1=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" dc:title "+'"{}" .'.format(row['Project Title'])
        type1=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" rdf:type "+"<http://purl.obolibrary.org/obo/OBI_0001636> ."
        pi=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000171 "+'"{}" .'.format(row['Contact PI / Project Leader'])
        university=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000174 " +'"{}" .'.format(row['Organization Name'])
        foa=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000172 "+'"{}" .'.format(row['FOA'])
        agency=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000173 "+'"{}" .'.format(row['Funding Organization'])
        start=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000169 "+'"{}" .' .format(row['Project Start Date'])
        end=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" obo:IDE_0000000170 "+'"{}" .'.format(row['Project End Date'])
        indexed=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" <http://purl.obolibrary.org/obo/APOLLO_SV_00000325>"+ '"{}" .'.format(today)
        linkout=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+"<http://purl.obolibrary.org/obo/ERO_0000480>"+" <http://www.ncbi.nlm.nih.gov/pubmed/{}> ".format("".join(row['pubid'][0]))+" ."
        indexed_by=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+" <http://www.pitt.edu/obc/IDE_0000003897> "+ '"{}" .'.format(email_id)
        about=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((auto_incrementNumber(number1))))+"  <http://purl.obolibrary.org/obo/IAO_0000136> <http://www.pitt.edu/obc/IDE_0000000229> ."
        number2=auto_incrementNumber(number1)
        predicate=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((number2+1)))+" obo:IAO_0000219 "+('<http://www.pitt.edu/obc/IDE_ARTICLE_{}> .'.format(str(auto_incrementNumber(number1))))
        grant=('<http://www.pitt.edu/obc/IDE_ARTICLE_{}>'.format((number2+1)))+" rdfs:label"+' "{}" .'.format(row['Project Number'])
        rdf.append(title1)
        rdf.append(type1)
        rdf.append(linkout)
        rdf.append(pi)
        rdf.append(university)
        rdf.append(foa)
        rdf.append(agency)
        rdf.append(start)
        rdf.append(end)
        rdf.append(indexed)
        rdf.append(indexed_by)
        rdf.append(about)
        rdf.append(predicate)
        rdf.append(grant)
        number1=number2+1
        filename=path+'\\'+'pub_grants_in_obc_no_dash.ttl'
        with open (filename, 'w') as fp:
            fp.write('@prefix rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .\n')
            fp.write('@prefix obo: <http://purl.obolibrary.org/obo/> .\n')
            fp.write('@prefix dc: <http://purl.org/dc/elements/1.1/> .\n')
            for line in rdf:
                fp.write(line + '\n')
                
if __name__ == '__main__':
    foa= input("enter Foa number")
    chromedriver_path= input("enter path of chrome driver")
    number1=int(input("enter number for indexing "))
    foa_number=getting_FOA(foa,chromedriver_path)
    file_path=download_excel(foa_number)
    df3,filename=grants_list(file_path)
    subprocess.call(['python',path+"\\"+'search1.py',filename] )
    #dict1=s1.generate_xml(filename)
    #time.sleep(20)
    dict1=o.get_xml()
    print(dict1)
    ttl_file(df3,number1,dict1)
    #subprocess.call(['python',path+"\\"+'obc.py'] )
