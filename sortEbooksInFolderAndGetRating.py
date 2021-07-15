#This program will find the names, athors, rating for the ebboks in a folder

import os
import xlsxwriter
import asyncio
import aiohttp

#input
loc = "//OMV2M\Publicmappe/bøker/751 Supense e-Books"
excelFileName = "751 Supense e-Books-test"  #to update just use the same index as before and it'll just uptate the new stuff. Althougth it will wil not remove books if  they are removed

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
stopAfterNumOfBooks = False
def createIndex(): #create an index so that if there is a need to update the file it does not have to update everything all agiain
    index = open(os.path.join(excelFileLoc, indexFileName + ".txt" ),"w+")
    index.write(str(books))
    index.close()


async def getWebpage(session, name):
    async with session.get("https://www.goodreads.com/search?utf8=%E2%9C%93&q=" + name + "&search_type=books") as resp:
        resp = await resp.read()
        if "<h1>page unavailable</h1>" in str(resp): #test for checing if goodreads reurns an almost ampty page
            #await asyncio.sleep(20)
            #resp = await resp.read()
            print(" <-||WARNING||->  Webpage is not loading...  The program is waiting for the delay to stop")
        return str(resp)

def getRating(page): #finds the rating in the html code to the "page"
    page = page[page.find("avg rating")-5:]
    page = page[:4]
    return page

def getGenre(page): #finds the genre in the html code to the "page"
    page = page[page.find("/shelf/show/"):page.rfind("/shelf/show/")+30]
    count = page.count("/shelf/show/")
    genre = ""

    #extracting the differnet genres from the jumbl
    for i in range(0, count):
        page = page[page.find("/shelf/show/")+12:]
        genre = genre + ", " + page[:page.find('"')]

    return genre[2:] #deleting the first comma and space

def getAuthor(page): #finds the athor in the html code to the "page"
    page = page[page.find('><span itemprop="name">'):]
    page = page[page.find('><span itemprop="name">')+6:]
    page = page[page.find(">")+1:page.find("<")]
    return page

def getLink(page): #finds the link to the book in the html code to the "page"
    page = page[page.find('<a class="bookTitle" itemprop="url" href="')+42:]
    page = "https://www.goodreads.com" + page[:page.find('"')]
    return page

def getName(page):
    page = page[page.find('<a class="bookTitle" itemprop="url" href="'):]
    page = page[page.find('>')+69:]
    page = page[:page.find("<")]
    return page

def deleteFirst(page):
    return page[page.find(getAuthor(page))+10:]



def containdigit(string):
    for character in string:
        if character.isdigit():
            return True
    return False
def findDigit(string):
    i = 0
    for character in string:
        if character.isdigit():
            return i
        i+=1
    return -1
def getAllFilesInDir(location):
    f = []
    allFiles = os.walk(loc, topdown=True)
    for (root, dirs, files) in allFiles:
        for file in files:
            if (file[file.rfind(".")+1:] in bannedFileTypes) == False:

                f.append(file)
                if len(f) >= stopAfterNumOfBooks and stopAfterNumOfBooks != False: #this is for testing
                    return f
    if f == []:
        print(" <-||ERROR||-> There are no files to walk")
    return f
def devideList(list, n):
    return [list[i:i + n] for i in range(0, len(list), n)]

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
    try:
        oneCharVal = 200/(sumList(charsInComp[1]) + sumList(charsInStr[1]))
        sameChars = 0
        for i in range(0,min([len(charsInStr[1]),len(charsInComp[1])])):
            sameChars += min([charsInStr[1][i],charsInComp[1][i]])
        return oneCharVal * sameChars
    except:
        print(" <-||ERROR||->  Something Went wrong in the function: howSimmilarLetters(",str, ", ", compareTo, ")")
        return -1

def compareWords(str,compareTo): # returns the precent of letters that are in the same position and are the same ut uses the length of the shortest word
    try:
        if compareTo in str:
            return 81
        lenShortestWord = len(min([str,compareTo], key=len))
        oneCharVal = 100/lenShortestWord
        percent = 0
        for i in range(0,lenShortestWord):
            if str[i] == compareTo[i]:
                percent += oneCharVal
        return percent
    except:
        return -1

def combineSeriesAndName(name, series, saga=""):
    if containdigit(series):
        if series.find(" Vol ") == -1:
            if findDigit(series) == series.find("0"):
                series= series[:series.find("0")] + ", #" + series[series.find("0"):]
            else:
                series = series[:findDigit(series)-1] +", #"+ series[:findDigit(series)]
        else:
            series = series.replace(" Vol ", " #")
    if saga != "":
        name = name + " ("+ saga + ": " + series + ")"
    else:
        name = name + " (" + series + ")"

    return name


#methods for getting soting the file name to name and author
def methodNameAuthor(header, differentiator = " - "): #works for 1 or more joints
    name = header[:header.rfind(differentiator)]
    author = header[header.rfind(differentiator)+3:]
    return name, author

def methodAuthorName(header, differentiator = " - "): #works for 1 or more joints
    name = header[header.rfind(differentiator)+3:]
    author = header[:header.rfind(differentiator)]
    return name, author

def methodSagaSeriesName(header, differentiator = " - "): # works with 2 or more, there can be bulshit between saga and series
    saga = header[:header.find(differentiator)]
    name = header[header.rfind(differentiator)+3:]
    series = header[:header.rfind(differentiator)]
    series = series[series.rfind(differentiator)+3:]
    if series != name:
        name = combineSeriesAndName(name,series, saga)
    return name, "404"

def methodSagaName(header, differentiator = " - "):
    saga = header[:header.find(differentiator)]
    name = header[header.rfind(differentiator)+3:]
    return saga + ": " + name, "404"

def methodAuthorSeriesName(header, differentiator = " - "): #works for 2 or more
    author = header[:header.find(differentiator)]
    name = header[header.rfind(differentiator)+3:]
    series = header[:header.rfind(differentiator)]
    series = header[header.rfind(differentiator)+3:]
    if series != name:
        name = combineSeriesAndName(book, series)
    return name, author

sortingMethods = [[methodNameAuthor,methodAuthorName],[methodAuthorSeriesName, methodSagaSeriesName, methodSagaName]]

async def tryMethods(session, header, methods):
    for method in methods:
        nameAndAuthor = method(header)
        page = await getWebpage(session, nameAndAuthor[0])
        if howSimilarLetters(getName(page), nameAndAuthor[0]) > 60 or compareWords(getName(page), nameAndAuthor[0]) > 80:
            if howSimilarLetters(getAuthor(page), nameAndAuthor[1]) > 60 or compareWords(getAuthor(page), nameAndAuthor[1]) > 80 or nameAndAuthor[1] == "404":
                return page
            else:
                for i in range(0,10):
                    if howSimilarLetters(getAuthor(page), nameAndAuthor[1]) > 60:
                        return page
                    page = deleteFirst(page)
    print(" <-||ERROR||-> There are no suiteble methods for this book: ", header)
    return "404"

async def newGetInfo(session, header):
    fullEntry = header
    format = header[header.rfind(".")+1:]
    header = header[:header.rfind(".")]

    numOfJoints = header.count(" - ")
    page = ""
    if numOfJoints == 0:
        print(" <-||ERROR||-> The header has no joints: ", header)
    elif numOfJoints == 1:
        page = await tryMethods(session, header, sortingMethods[0])
    elif numOfJoints == 2:
        page = await tryMethods(session, header, sortingMethods[1]+sortingMethods[0])
    else: #moe than three joins for example dude - book - series - the series of the series
        page = await tryMethods(session, header, sortingMethods[1]+sortingMethods[0]) #redundant, it could be moved But if some more methods are added then it should be there


    rating = getRating(page)
    genre = getGenre(page)
    link = getLink(page)
    name = getName(page)
    author = getAuthor(page)

    books.append((name, author, format, rating, genre, fullEntry, link))

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
        allFiles = devideList(getAllFilesInDir(loc), 50)
        for dList in allFiles:
            tasks = []

            for file in dList:
                if file in filesInLastBooks:
                    nr = filesInLastBooks.index(file)
                    allSortingMethods = [j for sub in sortingMethods for j in sub]
                    list = [howSimilarLetters(lastBooks[nr][1], allSortingMethods[i](file)[1]) for i in range(len(allSortingMethods))]
                    if any(i >= 60 for i in list):
                        books.append(lastBooks[nr])
                    else:
                        #add a delay between the tasks, for goodreads returns an almost empty page if the program is to fast
                        tasks.append(asyncio.ensure_future(newGetInfo(session, file)))
                else:
                    #add a delay between the tasks, for goodreads returns an almost empty page if the program is to fast
                    tasks.append(asyncio.ensure_future(newGetInfo(session, file)))
                if i > 25: #create a backup in the index, so the program can be stopped and resumed and not have to start over
                    createIndex()
                    i = 0
                else:
                    i += 1
            print("sleeping?")
            await asyncio.gather(*tasks)
            #await asyncio.sleep(1)
            print("Was sleeping")

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

#create an index so that if there is a need to update the file it does not have to update everything all agiain

createIndex()
