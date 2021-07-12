from io import BufferedRWPair
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import urllib.request
import re

#스크래핑 하기 
#엑셀 작성 준비 
wb = Workbook()
ws1 = wb.active
ws1.title = "유진이 숙제1 "
ws1.append(["Authers", "title","publisher","Abstract0","Abstract1","keywords","Pubdate","title+abstract","title+abstract+keywords","종류"])




#ws1.append(["Authers", "title","Abstract0","Abstract1","keywords","Pubdate"])
#1행부터 a-1 행까지인임 
#기본은 11이고 10행이다 

def work1(a):
    for i in range(1, a):
        print ("행번호:", i)
        addr ="//*[@id='divContent']/div[2]/div/div[3]/div[2]/ul/li[{}]/div[2]/p[1]/a".format(i)
    
        browser.find_element_by_xpath(addr).click()

        #browser.find_element_by_xpath("//*[@id='divContent']/div[2]/div/div[3]/div[2]/ul/li[1]/div[2]/p[1]/a").click()
        soup = BeautifulSoup(browser.page_source, 'html.parser')

        #제목
        Title= soup.find( "h3", attrs={"class":"title"})

        Title= Title.get_text()
        Title = Title.replace("\n", "")
        Title = Title.replace("\t", "")
        Title= Title.split('=')[0]

        abstract_tag= soup.findAll( "div", attrs={"class":"text off"})

        #요약 
        # print (abstract_tag)
        # print (len(abstract_tag))
        if len(abstract_tag) == 1:
            Abstract0=abstract_tag[0].get_text()
            Abstract1="없음"
        elif len(abstract_tag) == 0:
            Abstract0="없음"
            Abstract1="없음"
        else:
            Abstract0=abstract_tag[0].get_text()
            Abstract1=abstract_tag[1].get_text()
        Abstract0 = Abstract0.replace("\n", "")
        Abstract1 = Abstract1.replace("\n", "")
        #print(Abstract)

        #Authers
        Authers=soup.find( "div", attrs={"class": "infoDetailL"}).find("li").find("p").get_text()
        Authers = Authers.replace("\n", "")
        Authers = Authers.replace("\t", "")

        #학술지명 
        imsi1=soup.find( "div", attrs={"class": "infoDetailL"}).findAll("li")
        publisher=imsi1[2].find("p").get_text()
        publisher = publisher.replace("\n", "")
        publisher = publisher.replace("\t", "")

        print (publisher)
        #년도구하기
        #imsi1=soup.find( "div", attrs={"class": "infoDetailL"}).findAll("li")
        pubdate=imsi1[4].find("p").get_text()

        keywords=imsi1[6].find("p").get_text()
        keywords = keywords.replace("\n", "")
        keywords = keywords.replace("\t", "")

        # Authers= Authers.encode('utf-8').strip()
        # Title= Title.encode('utf-8').strip()
        # Abstract0= Abstract0.encode('utf-8').strip()
        # Abstract1= Abstract1.encode('utf-8').strip()
        #print(type(Abstract1))
        #Abstract1 = str (Abstract1)
        #Abstract1= Abstract1.encode('utf-8').strip()
        # pubdate= pubdate.encode('utf-8').strip()
        Authers = ILLEGAL_CHARACTERS_RE.sub(r'',Authers)
        Title = ILLEGAL_CHARACTERS_RE.sub(r'',Title)

        Abstract0 = ILLEGAL_CHARACTERS_RE.sub(r'',Abstract0)
        Abstract1 = ILLEGAL_CHARACTERS_RE.sub(r'',Abstract1)
        keywords = ILLEGAL_CHARACTERS_RE.sub(r'',keywords)
        pubdate = ILLEGAL_CHARACTERS_RE.sub(r'',pubdate)
        
        ws1.append([Authers, Title,publisher,Abstract0,Abstract1,keywords,pubdate])
        #print(Abstract1)
        #ws1.append([Abstract1])

        browser.back()     


#################### START ################       

#크롬 열기 
browser = webdriver.Chrome()

browser.get("http://www.riss.kr.eproxy.sejong.ac.kr/index.do")

#로그인하기 
elem=browser.find_element_by_xpath("//*[@id='userID']")
elem.send_keys("")

elem=browser.find_element_by_xpath("//*[@id='password']")
elem.send_keys("")

elem=browser.find_element_by_xpath("//*[@id='loginForm']/dd[2]/a/img")

elem.click()

time.sleep(3)

#검색하기 
elem=browser.find_element_by_xpath("//*[@id='query']")
# elem=browser.find_element_by_id("query")
#print (browser.a)

elem.send_keys("(중도입국학습자|중도입국자녀|중도입국청소년) (한국어교육)")
elem.send_keys(Keys.ENTER)

#국내학술논문 클릭 
browser.find_element_by_xpath("//*[@id='divContent']/div/div/div[2]/div[1]/div[2]/a/img").click()

time.sleep(3)




print("시작")
#첫 페이지 실행
work1(11)

#2번째 페이지 수 지정
#3부터 2번째 페이지 (-1씩) 
# for i in range(3, 12):
for i in range(3, 9):
    print ("페이지:", i-1)
    addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
    browser.find_element_by_xpath(addr).click()

    work1()

# i = 12
# print ("첫번쨰 넥스트 ")
# addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
# browser.find_element_by_xpath(addr).click()

# work1()

# for i in range(4, 13):
#     print ("페이지:", i-1)
#     addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
#     browser.find_element_by_xpath(addr).click()

#     work1()

# i = 13
# print ("두번쨰 넥스트 ")
# addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
# browser.find_element_by_xpath(addr).click()
# work1()

# for i in range(4, 7):
#     print ("페이지:", i-1)
#     addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
#     browser.find_element_by_xpath(addr).click()

#     work1()

# #8 -> 7로 바꿔야 되는거 아닌가??
# #마지막 페이지에는 행수가 다름. 

# wb.save(filename='유진이숙제1_1.xlsx')

# i = 7
# print ("마지막 페이지:")
# addr ="//*[@id='divContent']/div[2]/div/div[4]/a[{}]".format(i)
# browser.find_element_by_xpath(addr).click()

# for i in range(1, 3):
#     print ("행번호:", i)
#     addr ="//*[@id='divContent']/div[2]/div/div[3]/div[2]/ul/li[{}]/div[2]/p[1]".format(i)

#     browser.find_element_by_xpath(addr).click()

#     #browser.find_element_by_xpath("//*[@id='divContent']/div[2]/div/div[3]/div[2]/ul/li[1]/div[2]/p[1]/a").click()
#     soup = BeautifulSoup(browser.page_source, 'html.parser')

#     #제목
#     Title= soup.find( "h3", attrs={"class":"title"}).get_text()
#     Title = Title.replace("\n", "")
#     Title = Title.replace("\t", "")
#     Title= Title.split('=')[0]

#     abstract_tag= soup.findAll( "div", attrs={"class":"text off"})

#     #요약 
#     # print (abstract_tag)
#     # print (len(abstract_tag))
#     if len(abstract_tag) == 1:
#         Abstract0=abstract_tag[0].get_text()
#         Abstract1="없음"
#     elif len(abstract_tag) == 0:
#         Abstract0="없음"
#         Abstract1="없음"
#     else:
#         Abstract0=abstract_tag[0].get_text()
#         Abstract1=abstract_tag[1].get_text()
#     Abstract0 = Abstract0.replace("\n", "")
#     Abstract1 = Abstract1.replace("\n", "")
#     #print(Abstract)

#     #Authers
#     Authers=soup.find( "div", attrs={"class": "infoDetailL"}).find("li").find("p").get_text()
#     Authers = Authers.replace("\n", "")
#     Authers = Authers.replace("\t", "")

#     #년도구하기
#     imsi1=soup.find( "div", attrs={"class": "infoDetailL"}).findAll("li")
#     pubdate=imsi1[4].find("p").get_text()

#     keywords=imsi1[6].find("p").get_text()
#     keywords = keywords.replace("\n", "")
#     keywords = keywords.replace("\t", "")

#     # Authers= Authers.encode('utf-8').strip()
#     # Title= Title.encode('utf-8').strip()
#     # Abstract0= Abstract0.encode('utf-8').strip()
#     # Abstract1= Abstract1.encode('utf-8').strip()
#     #print(type(Abstract1))
#     #Abstract1 = str (Abstract1)
#     #Abstract1= Abstract1.encode('utf-8').strip()
#     # pubdate= pubdate.encode('utf-8').strip()
#     Authers = ILLEGAL_CHARACTERS_RE.sub(r'',Authers)
#     Title = ILLEGAL_CHARACTERS_RE.sub(r'',Title)

#     Abstract0 = ILLEGAL_CHARACTERS_RE.sub(r'',Abstract0)
#     Abstract1 = ILLEGAL_CHARACTERS_RE.sub(r'',Abstract1)
#     keywords = ILLEGAL_CHARACTERS_RE.sub(r'',keywords)
#     pubdate = ILLEGAL_CHARACTERS_RE.sub(r'',pubdate)
    
#     ws1.append([Authers, Title,Abstract0,Abstract1,keywords,pubdate])
#     #print(Abstract1)
#     #ws1.append([Abstract1])




    #파일이름
print("종료")
browser.quit()
wb.save(filename='유진이숙제1_2.xlsx')
