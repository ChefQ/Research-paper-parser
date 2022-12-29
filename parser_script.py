
from selenium import webdriver 
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import PyPDF2.utils
#import PyPDF2
import requests
import regex as re
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select
import vlc
import pdfplumber
import regex as re
import pdb
import regex as re
from PyPDF2 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import pytesseract
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SAMPLE_SPREADSHEET_ID = '1kge8VOTe7oUNFagXmP2f3HdgI0j4-y2wOQlcTjuuTy0'
SAMPLE_RANGE_NAME = 'A2:D7'
p = vlc.MediaPlayer("/Users/oluwaseuncardoso/PythonStuff/AlarmNuclearMeltd HYP013301_preview.mp3")


def findLink(tag):
    if tag.name == "span" and tag.a != None and tag.get("class") != None:
        cond = "gateway.proquest.com"
        if tag["class"][0] == 'subjectField-postProcessingHook' and cond in tag.a["href"]:
            return True
    return False
def findCapcha(tag):
    if tag.name =="div" and div.get("id") == "start":
        if tag.div.get("class") == "alert alert-info container captcha_alert_box":
            return True
    return False

def getBs(link):
    """
    This function takes a proquest Url
    and returns the beautiful soup object
    """
    driver.get(link)
    sleep(3)
    html = driver.execute_script("return document.documentElement.outerHTML;")
    bs = BeautifulSoup(html, "html5lib")
    capcha1 = bs.find_all("div", attrs = {"class": "alert alert-info container captcha_alert_box"} )
    playing = False
    while(len(capcha1) != 0 ):
        print("start playing alarm")
        p.play()
        playing = True
        sleep(5)
        break
    if(playing == True):
        p.stop()
    return bs

def goToPage(num): 
    """
    This function takes the browser to the specific page num
    This fucntion is useful because sometimes mid way in downloading.
    A bot check stops the program.
    We need a way to go back to our intended page number to continue downloading
    """
    pass

def writeContent(pdf): # come here to change code
    response = requests.get(pdf) 
    with open('./dissertation.pdf', 'wb') as f:
        f.write(response.content)
        sleep(4)  
    ack1 = findAck()
    ack2 = findAck2()
    return ack1, ack2

def getAckFromLines(lines,ACK2=None):
    # numeral numbers
    regex1 = r"[^\w]+i$+|^i$|^1[^\w]+|[^\w]+1$+|^1$|"
    regex2 = r"^ii[^\w]+|[^\w]+ii$+|^ii$|^2[^\w]+|[^\w]+2$+|^2$|"
    regex3 = r"^iii[^\w]+|[^\w]+iii$+|^iii$|^3[^\w]+|[^\w]+3$+|^3$|"
    regex4 = r"^iv[^\w]+|[^\w]+iv$+|^iv$|^4[^\w]+|[^\w]+4$+|^4$|"
    regex5 = r"^v[^\w]+|[^\w]+v$+|^v$|^5[^\w]+|[^\w]+5$+|^5$|"
    regex6 = r"^vi[^\w]+|[^\w]+vi$+|^vi$|^6[^\w]+|[^\w]+6$+|^6$|"
    regex7 = r"^vii[^\w]+|[^\w]+vii$+|^vii$|^7[^\w]+|[^\w]+7$+|^7$|"
    regex8 = r"^viii[^\w]+|[^\w]+viii$+|^viii$|^8[^\w]+|[^\w]+8$+|^8$|"
    regex9 = r"^ix[^\w]+|[^\w]+ix$+|^ix$|^9[^\w]+|[^\w]+9$+|^9$|"
    regex10 = r"^x[^\w]+|[^\w]+x$+|^x$|^10[^\w]+|[^\w]+10$+|^10$"
    regex = regex1+regex2+regex3+regex4+regex5+regex6+regex7+regex8+regex9+regex10
    
    ack_index = -1
    if ACK2 != None:
        for i in range(len(lines)):
            if re.search(ACK2, lines[i],  re.IGNORECASE):
                ack_index = i
                break
        lines = lines[ack_index+1:]

    # remove every after page number(if page number exists)
    end_line = len(lines) 
    half_page_line = len(lines)//2 # the page number will must likes be between the bottom and half the page
    last_page_line = len(lines)-1
    for i in range(last_page_line, half_page_line, -1):
        if re.search(regex, lines[i]):
            #pdb.set_trace()
            end_line = i
            break
    return lines[:end_line]
    
def findAck(pdf_path = "./dissertation.pdf",from_page = 0,to_page = 29):
    """
    create a new pdf file from a subsection from pdf
    from_page(int): Where to start. Starts from 0.
    to_page(int): Where to end(inclusive).
    """
    pdf = PdfFileReader(pdf_path)
    num_pages = pdf.getNumPages()
    if to_page > num_pages:
        to_page = num_pages-1
    ack_found = None
    ack_extracted, content_extracted = False, False
    acknowledgment = "Acknoledgement not present in file"
    for page_num in range(from_page, to_page):
        try: 
            pdfWriter = PdfFileWriter()
            pdfWriter.addPage(pdf.getPage(page_num))   
            with open(f'./temp.pdf', 'wb') as temp:
                pdfWriter.write(temp)
            temp.close()
            content = convertToString().strip()
        except PyPDF2.utils.PdfReadError:
            return {"bool" : False, "content" :"PDF Broken"}
        if(ack_found == None):
            ACK = r'ack[n]?owledg[e]?ment[s]?\s*\n'
            ack_found = re.search(ACK,content, re.IGNORECASE)
        if(ack_found != None and ack_extracted == False):
            #pdb.set_trace()
            ACK2 = ack_found[0].split("\n")[0] 
            #find the index to find the index in lines
            
            try:
                ack_index = content.split("\n").index(ACK2)
            except ValueError:
                return {"bool" : False , "content" : "Error: Special case please review"}
            
            if ack_index > 4:
                ack_found = None
            else:
                acknowledgment =  content.strip()     
                lines = acknowledgment.split('\n')          
                lines = getAckFromLines(lines,ACK2)
                acknowledgment = "\n".join(lines)
                ack_extracted  = True
            continue # start loop again
        if(ack_extracted == True and content_extracted == False):
            #pdb.set_trace()
            lines = content.split('\n')
            next_title = lines[ack_index]
            words_of_next_title = next_title.split(" ")
            words_of_next_title = [i for i in words_of_next_title if i not in [""," "]]
            #pdb.set_trace()
            there_is_no_title = len(words_of_next_title) > 6
            if (there_is_no_title):   
                lines = content.strip().split('\n')          
                lines = getAckFromLines(lines)
                acknowledgment += "\n".join(lines)
                last_word = lines[-1]                    
            else:
                content_extracted = True
        if content_extracted == True:            
            return {"bool" : True, "content" :acknowledgment}
        
       
    return {"bool" : False, "content" :"ACKNOTFOUND"}

def convertToString():
    images = convert_from_path("./temp.pdf")
    content = ""
    images[0].save(f"./temp.jpg","JPEG")
    content = pytesseract.image_to_string(f"./temp.jpg")
    return content
import numpy as np
def Levenshtein(r, h):
    """                                                                         
    Calculation of WER with Levenshtein distance.                               
                                                                                
    Works only for iterables up to 254 elements (uint8).                        
    O(nm) time ans space complexity.                                            
                                                                                
    Parameters                                                                  
    ----------                                                                  
    r : list of strings                                                                    
    h : list of strings                                                                   
                                                                                
    Returns                                                                     
    -------                                                                     
    (WER, nS, nI, nD): (float, int, int, int) WER, number of substitutions, insertions, and deletions respectively
                                                                                
    Examples                                                                    
    --------                                                                    
    >>> wer("who is there".split(), "is there".split())                         
    0.333 0 0 1                                                                           
    >>> wer("who is there".split(), "".split())                                 
    1.0 0 0 3                                                                           
    >>> wer("".split(), "who is there".split())      #ask in Pia
    Inf 0 3 0                                                                           
    """

    n = len(r) # The number of words in REF
    m = len(h) # The number of words in HYP
    R = np.zeros((n+1,m+1))
    B = np.zeros((n+1,m+1))

    #for all i,j s.t.  i = 0 or  j = 0,	set	R[i,j] ← max (i,j) end
    R[0,:] = np.arange(m+1)
    R[:,0] = np.arange(n+1)
    # i think we should do this aswell
    up = 0
    left = 1
    up_left = 2
    up_left2 = 3
    B[0,:] = left
    B[:,0] = up
    B[0,0] = up
    for i in range(1,n+1):
        for j in range(1,m+1):
            dele = R[i - 1, j] + 1 # delete
            sub = R[i - 1, j - 1] + (1,0)[r[i-1] == h[j-1]] #substitute #NOTE look at this
            ins = R[i, j-1] + 1 #insert

            R[i,j] = min(dele,sub,ins)
            if R[i,j] == dele:
                B[i , j] = up
            elif R[i , j] == ins:
                B[i,j] = left
            else:
                B[i,j] = (up_left, up_left2 )[r[i-1] == h[j-1]]
    i,j = n,m
    nSub,nDel,nIns = 0, 0, 0
    transversal = True
    while transversal == True:
        path = B[i,j]
        if i <=  0 and j <=0:
            transversal = False
            break
        if path == up_left:
            i -= 1
            j -= 1
            nSub += 1
        elif path == left:
            j -= 1
            nIns +=1
        elif path == up:
            i -= 1
            nDel +=1
        else: # correct
            i -= 1
            j -= 1
    return R[n,m]/n, nSub, nIns, nDel        
    
def getAcknoledgement(page,pages):
    """
    Rule:
    1. Look for the first occurcance of Acknowledgments and a new line.
    2. Caputure the content until you hit the title of the next section  
    #ASSUMPTIONS that the dissertations need to follow.
    #1. Every dissertation starts it's acknoledgement with the word "Acknoledgement" as the only word in it's title
    #2. No two titles can occupy the same page. 
    #3. The acknoledgement starts with a capital letter (i.e Acknoledgement)
    #4. The title acknoledgemen may or may not end an s.
    #5. If a line has 1 to 6 words in it. Then it's a title of a new section.
    #6. Titles are located on the same location on each page.i.e top and centre
    #7. Each word is of average size(6 chars long)
    #8. Target titles can either be Acknowledgments or acknowledg(e)ments
    #9. The Acknoledgemts page doesn't have picture/images
    #10. Acknoledgements are never in the first page
    Params: 
    page : pdfplumber.page.Page
    returns(String) :The content of the abstraction
    """    
    initial_index = page.page_number -1
    content = page.extract_text()
    if(content is None): return None
    content = content.strip()
    ACK = r'ack[n]?owledg[e]?ment[s]?\s*\n'
    ack_found = re.search(ACK,content, re.IGNORECASE)
    title_len = len(ACK)
    acknowledgment = None
    #if Acknowledgement is the title extract it else continue
    if(ack_found):
        #pdb.set_trace()
        ACK2 = ack_found[0].split("\n")[0] 
        acknowledgment =  content.strip()     
        lines = acknowledgment.split('\n')          
        lines = getAckFromLines(lines,ACK2)
        acknowledgment = "\n".join(lines)        
        stop_not_true = True
        #find the index to find the index in lines
        try:
            ack_line = content.strip().split("\n")
            ack_index = ack_line.index(ACK2)
            ack_line = [ack_line[i] for i in range(ack_index+1) if i not in ["", " "]]
            ack_index = ack_line.index(ACK2)
        except ValueError:
            return "Error: Special case please revuew"
        counter = 1
        while(stop_not_true):
            next_page = counter + initial_index
            #pdb.set_trace()
            page = pages[next_page]
            content = page.extract_text().lower().lstrip()
            lines = content.split('\n')
            next_title = lines[ack_index]
            words_of_next_title = next_title.split(" ")
            words_of_next_title = [i for i in words_of_next_title if i not in [""," "]]
            #pdb.set_trace()
            there_is_no_title = len(words_of_next_title) > 6
            if (there_is_no_title):   
                lines = content.strip().split('\n')          
                lines = getAckFromLines(lines)
                acknowledgment += "\n".join(lines)
                last_word = lines[-1]                    
                counter+=1
            else:
                stop_not_true = False
    return acknowledgment

def findAck2(pdf_path = "./dissertation.pdf",from_page = 0,to_page = 29):
    with pdfplumber.open(pdf_path, strict_metadata=True) as pdf:
        pages = iter(pdf.pages)
        count = 0
        
        for page in pages:
            count +=1
            if count >= from_page: 
                content = page.extract_text()        
                acknowledgement = getAcknoledgement(page, pdf.pages)    
                if(acknowledgement != None):
                    if("(cid:" in acknowledgement):
                        acknowledgement = "ERROR: It contains embedded fonts. This requires revisitation."
                        
                    return {"bool": True ,  "content": acknowledgement}
            if page.page_number == to_page:
                print("ack not found")
                return {"bool": False, "content":"ACKNOTFOUND"  }
            

def compare(ack1,ack2):
    if(ack1['bool'] == True and ack2['bool']  == True ):
        content1 = ack1["content"].replace("\n", "").strip().split()
        content2 = ack2["content"].replace("\n", "").strip().split()
        WER = Levenshtein(content1, content2)[0]
        print(WER)
        if(WER < 0.2):
            return ack1["content"], "N/A: same as 1st parser"          
    return ack1["content"] , ack2["content"]


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())            
    return creds
creds = main()
service = build('sheets', 'v4', credentials=creds)

def send(data, cells, service):
    
    value_range_body = {
        "majorDimension" : "ROWS",
        "values" : 
            [
            [data['author'],data['title'], data['ack1'],data['ack2']],
            ]
    }

    # have a program that send s batch as oppossed to just one
    sheet = service.spreadsheets()
    request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                          range=cells,
                          valueInputOption ='RAW',
                          includeValuesInResponse = True,
                          body = value_range_body )
    try:
        return request.execute()
    except ConnectionResetError:
        print("ConnectionResetError")
        creds = main()
        service = build('sheets', 'v4', credentials=creds)
        return send(data, cells, service)




s = Service("/Users/oluwaseuncardoso/Downloads/chromedriver")
driver = webdriver.Chrome(service = s)
driver.get("https://www.proquest.com/shibboleth?accountid=14771")
sleep(3)    

user_name = "alebiosu"
pass_word = "Toluwanimi01"
username = driver.find_element_by_id("username") 
username.clear()
username.send_keys(user_name)
password = driver.find_element_by_id("password") 
password.clear()
password.send_keys(pass_word)
password.send_keys(Keys.RETURN)
sleep(4)

print("What page do you want to go to: ")
page = int(input())

print("How many pages per page do you want to see 10, 20, 50, 100: ")
items_per_page = int(input())

print(f"what orderd pdf do you want to start from. Pick 1 - {items_per_page} ")
start = int(input())
  
try:
    elem = driver.find_element_by_id("queryTermField")  
except  NoSuchElementException:
    elem = driver.find_element_by_id("searchTerm")
elem.clear()
elem.send_keys("STYPE(DISSERTATION) AND PD(2016) AND DEP.X(PSYCHOLOGY) AND LA(ENGLISH) AND DG(PHD) AND (ULO(CANADA) OR ULO(UNITED STATES))  NOT SU(CLINICAL)")
elem.send_keys(Keys.RETURN)
sleep(4)


dropdown = Select(driver.find_element_by_id("itemsPerPage"))
dropdown.select_by_value(str(items_per_page))


search = driver.find_element_by_id("pageNbrField") 
search.clear()
search.send_keys(str(page))
search.send_keys(Keys.RETURN)



html = driver.execute_script("return document.documentElement.outerHTML;")
bs = BeautifulSoup(html, "html5lib")
results = bs.find_all(attrs= {"class": "resultItem ltr"})
h1 = bs.find_all( "h1")[0]
num_results = h1.contents[0]
num_results = num_results.strip()
num_results = num_results.replace(',', "")
num_results = re.search(r'[0-9]*', num_results)
num_results = int(num_results[0])
next_page = bs.find_all("a", attrs = {"title" : "Next Page"} )
count = (page -1) * items_per_page + start



import csv
Error = "Does not exist in this dissertation"

while count <= num_results:
    print(f"starting downloads at page:{page}")
    print("---------------------------")

    for index in range(start-1, len(results)):
        data = {'ack1':Error, 'ack2': Error}
        cell = f"A{count+1}:D{count+1}"
        result = results[index]
        link = result.find(attrs = {"contentArea"}).find("a")["href"]
        bs = getBs(link)    
        div = result.find_all("div", attrs= {"class": "truncatedResultsTitle"})[0]
        skip = """
        if(len(div.contents) > 1):
            data['title'] = ""
            for content in div.contents:
                if(len(content.contents) > 1):
                    data['title'] += "".join([str(i) for i in content.contents])
                else:
                    data['title'] += "".join(content.contents)
        else:
            data['title'] = div.contents[0]
        """
        data['title'] = "".join([str(i) for i in bs.find_all("h1", attrs= {"class": "documentTitle"})[0].contents])
        # find author
        spans = result.find_all("span", attrs= {"class": "titleAuthorETC"})
        data['author'] = spans[0].contents[0].replace(".\xa0\n","")
        try:
            pdf = bs.find(attrs = {"download":"ProQuestDocument.pdf"})["href"]
            ack1, ack2 = writeContent(pdf)
            data['ack1'], data['ack2'] = compare(ack1,ack2)
            send(data,cell,service)
            
        except TypeError:           
            span = bs.find_all(lambda tag:findLink(tag))
            if(len(span) != 0):
                link = span.a["href"]
                bs = getBs(link) 
                pdf = bs.find(attrs = {"download":"ProQuestDocument.pdf"})["href"]
                ack1, ack2 = writeContent(pdf) 
                data['ack1'], data['ack2'] = compare(ack1,ack2)
                send(data,cell,service)
            else:
                print("Couln't download PDF")
                # find title
                send(data,cell,service)
        title = data['title']         
        print(f'pdf: {count}, cell: {cell} title:{title} ') 
        count +=1
               
    if len(next_page) == 1: # there is a next page
        page+=1
        print(f'going to next page')
        next_page_link = next_page[0]['href']
        bs = getBs(next_page_link)
        results = bs.find_all(attrs= {"class": "resultItem ltr"}) 
    next_page = bs.find_all("a", attrs = {"title" : "Next Page"} )
    start = 1


count
start = count - (page-1)*items_per_page 

count,start


def getAffiliate(cells):
    result = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=cells).execute()
    rows = result.get('values', [])
    rows = [i[0] for i in rows]
    return rows


import csv
Error = "Does not exist in this dissertation"
affiliations, authors = [], []
while count <= num_results:
    print(f"starting downloads at page:{page}")
    print("---------------------------")
    for index in range(start-1, len(results)):
        data = {'ack1':Error, 'ack2': Error}
        cell = f"A{count+1}:D{count+1}"
        result = results[index]      
        
        spans = result.find_all("span", attrs= {"class": "titleAuthorETC"})
        author = spans[0].contents[0].replace(".\xa0\n","")
        affiliation = spans[1].contents[0]
        content = re.search(r'[a-zA-Z\s-,()]*',affiliation)
        affiliation = content[0]
        authors.append(author)
        affiliations.append(affiliation)
        count +=1
    if len(next_page) == 1: # there is a next page
        page+=1
        print(f'going to next page')
        next_page_link = next_page[0]['href']
        bs = getBs(next_page_link)
        results = bs.find_all(attrs= {"class": "resultItem ltr"}) 
    next_page = bs.find_all("a", attrs = {"title" : "Next Page"} )
    start = 1



    
authors_g = getAffiliate("A2:1159")


len(affiliations)


authors_g == authors


value_range_body = {
    "majorDimension" : "COLUMNS",
    "values" : 
        [
        affiliations,
        ]
}

# have a program that send s batch as oppossed to just one
sheet = service.spreadsheets()
request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                      range="E2",
                      valueInputOption ='RAW',
                      includeValuesInResponse = True,
                      body = value_range_body )


result = request.execute()




