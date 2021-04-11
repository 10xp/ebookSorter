


#This program will find the names, athors, rating for the ebboks in a folder

import os
import xlsxwriter
import asyncio
import aiohttp

#input
loc = "//OMV2M\Publicmappe/bøker/Kindle Library 12-26-10/Library"
excelFileName = "test"  #to update just use the same index as before and it'll just uptate the new stuff. Althougth it will wil not remove books if  they are removed

excelFileLoc = loc[:loc.rfind("/")] #just a default, it will go in the folder of the folder of the excelFile and the index
indexFileName = excelFileName + "-index"  #just a default

#excel-variables
workbook = xlsxwriter.Workbook(excelFileLoc + "/" + excelFileName +'.xlsx') #loc[:loc.rfind("/")] + excelFileName +'.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 0


books = []

#things that need updating, if the program is changed
bannedFileTypes = {"jpg", "opf", "db", "tmp", "tmp-journal"}

#for testing use False if there is no limit
stopAfterNumOfBooks = 100
def createIndex():
    #create an index so that if there is a need to update the file it does not have to update everything alla agiain

    index = open(os.path.join(excelFileLoc, indexFileName + ".txt" ),"w+")
    index.write(str(books))
    index.close()

async def getWebpage(session, name):
    async with session.get("https://www.goodreads.com/search?utf8=%E2%9C%93&q=" + name + "&search_type=books") as resp:
        resp = await resp.read()
        return str(resp)

def getRating(page): #finds the rating in the html code to the "page"
    page = page[page.find("avg rating")-5:]
    page = page[:4]
    return(page)

def getGenre(page): #finds the genre in the html code to the "page"
    page = page[page.find("/shelf/show/"):page.rfind("/shelf/show/")+30]
    count = page.count("/shelf/show/")
    genre = ""

    #extracting the differnet genres from the jumbl
    for i in range(0, count):
        page = page[page.find("/shelf/show/")+12:]
        genre = genre + ", " + page[:page.find('"')]

    return(genre[2:]) #deleting the first comma and space

def getAuthor(page): #finds the athor in the html code to the "page"
    page = page[page.find('><span itemprop="name">'):]
    page = page[page.find('><span itemprop="name">')+6:]
    page = page[page.find(">")+1:page.find("<")]
    return(page)

def getLink(page): #finds the link to the book in the html code to the "page"
    page = page[page.find('<a class="bookTitle" itemprop="url" href="')+42:]
    page = page[:page.find('"')]
    return("https://www.goodreads.com"+page)

def getName(page):
    page = page[page.find('<a class="bookTitle" itemprop="url" href="'):]
    page = page[page.find('>')+69:]
    page = page[:page.find("<")]
    return(page)

def deleteFirst(page):
    return page[page.find(getLink(page)):]



def containdigit(string):
    for character in string:
        if character.isdigit():
            return(True)
    return(False)
def findDigit(string):
    i = 0
    for character in string:
        if character.isdigit():
            return(i)
        i+=1
    return(-1)
def getAllFilesInDir(location):
    f = []
    allFiles = os.walk(loc, topdown=True)
    for (root, dirs, files) in allFiles:
        for file in files:
            if (file[file.rfind(".")+1:] in bannedFileTypes) == False:
                f.append(file)
                if len(f) > stopAfterNumOfBooks and stopAfterNumOfBooks != False: #this is for testing
                    return(f)
    return(f)


def sumList(list):
    tot = 0
    for i in range(0,len(list)):
        tot += list[i]
    return tot
def numOfChar(str, letters=[]):  #returns  the differnet letters used and how many times they where used
    charsInStr = [0 for i in range(0,len(letters))]
    for character in str:
        if (character in letters) == False:
            letters.append(character)
            charsInStr.append(1)
        else:
            charsInStr[letters.index(character)] += 1
    return letters, charsInStr
def howSimilarLetters(str,compareTo): #returns the percent of lertters that are the same
    charsInStr = numOfChar(str)
    charsInComp = numOfChar(compareTo, charsInStr[0])
    chars = charsInComp[0]

    oneCharVal = 200/(sumList(charsInComp[1]) + sumList(charsInStr[1]))
    sameChars = 0
    for i in range(0,min([len(charsInStr[1]),len(charsInComp[1])])):
        sameChars += min([charsInStr[1][i],charsInComp[1][i]])
    return oneCharVal * sameChars


#methods for getting soting the file name to name and author

def nameAuthor(header):
    name = header[:header.rfind(" - ")]
    author = header[header.rfind(" - ")+3:]
    return name, author

def authorName(header):
    author = header[:header.find(" - ")]
    book = header[header.rfind(" - ")+3:]
    series = header[header.find("-"):]
    series = series[:series.find("-")]

    name = book
    if containdigit(series):
        if series.find("0") != -1:
            series.replace(" 0", ", #")
        else:
            series = series[:findDigit(series)-1] +", "+"#"+ series[:findDigit(series)]
        name = name + " (" + series + ")"

    name.replace(" -",":")
    return name, author

sortingMethods = [nameAuthor, authorName] #this is the differnet methods for getting the name and author so if one fails it's posseble to use a redundant method


def freshUpNameAuthor(name, author):
    name = name.replace(", the", "").replace("_",":")
    author = author.replace(";", ",")
    return name, author


async def getInfo(session, header):

    defaultMethod = 0

    fullEntry = header
    header = header.replace("_ a novel", "")

    format = header[header.rfind(".")+1:]
    header = header[:header.rfind(".")]
    rating = ""

    i = 0
    while rating.find(".") == -1 and i < len(sortingMethods):
        i+=1
        temp = sortingMethods[defaultMethod](header)

        name = temp[0]
        author = temp[1]

        #this uses the website goodreads to get the data
        page = await getWebpage(session, name)

        if howSimilarLetters(getAuthor(page), author) < 60:
            print("This book is probably wrong: ", name," --> ", author, " vs ", getAuthor(page))
            if howSimilarLetters(getAuthor(page), name) > 60:  #this will never come true (almost), so could be moved
                name = temp[1]
                author = temp[0]
                page = await getWebpage(session, name)
                print("name and author is swiched: ", name, author)
            else: #here i should try to remove
                print("Move througth the different books (use )", )
                pass
        else:
            print("Probably right book: ", name, " --> ", author, " vs ", getAuthor(page))
        rating = getRating(page)
        genre = getGenre(page)
        link = getLink(page)

        #its nessecary to have a new method for choosing the method, because this only supports two...
        if rating.find(".") == -1: #choosing the next method... in a very bad way...
            if defaultMethod == 1:
                defaultMethod = 0
            else:
                defaultMethod = 1
    if getName(page) != "":
        name = getName(page)
        author = getAuthor(page)
    book = (name, author, format, rating, genre, fullEntry, link)
    books.append(book)
    return



async def main():
    #this part of the program gets the indexFile
    filesInLastBooks = []
    pathIndex = excelFileLoc + "/" + indexFileName + ".txt"
    if os.path.isfile(pathIndex) and open(pathIndex, "r").read() != "":
        lastBooks = eval(open(pathIndex, "r").read()) #gets the index file
        for i in range(0,len(lastBooks)):
            filesInLastBooks.append(lastBooks[i][5])

    i = 0
    async with aiohttp.ClientSession() as session:
        tasks = []
        for file in getAllFilesInDir(loc):
            if file in filesInLastBooks:
                nr = filesInLastBooks.index(file)
                list = [howSimilarLetters(lastBooks[nr][1], sortingMethods[i](file)[1]) for i in range(len(sortingMethods))]
                if any(i >= 60 for i in list):
                    print("denne boka er antagelig riktig:", file, lastBooks[nr][1])
                    books.append(lastBooks[nr])
                else:
                    tasks.append(asyncio.ensure_future(getInfo(session, file)))
            else:
                tasks.append(asyncio.ensure_future(getInfo(session, file)))
            if i > 25: #create a backup in the index, so the program can be stopped and resumed and not have to start over
                createIndex()
                i = 0
            else:
                i += 1
        await asyncio.gather(*tasks)

loop = asyncio.get_event_loop()
loop.run_until_complete(main())

#create a xcel document
tableSize = 'A1:F'+ str(len(books)+1)
worksheet.add_table(tableSize, {'data': books, "columns": [{'header':"Book"}, {'header':"Author(s)"}, {'header':"Format"}, {'header':"Rating"}, {'header':"Genre"}, {'header':"fullFileName"}]})
for i in range(0, len(books)): #makes the name a link to get more information
    worksheet.write_url(row, 0, books[i][6], string=books[i][0])
    row+=1
    pass
workbook.close()

#create an index so that if there is a need to update the file it does not have to update everything alla agiain

createIndex()
