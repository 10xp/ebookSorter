


#This program will find the names, athors, rating for the ebboks in a folder

import os
import requests
import xlsxwriter

#input
loc = "//OMV2M/Publicmappe/bøker/Kindle Library 12-26-10/Library"
excelFileName = "test"  #to update just use the same index as before and it'll just uptate the new stuff. Althougth it will wil not remove books if  they are removed

excelFileLoc = loc[:loc.rfind("/")] #just a default, it will go in the folder of the folder of the excelFile and the index
indexFileName = excelFileName + "-index"  #just a default

#excel-variables
workbook = xlsxwriter.Workbook(excelFileLoc + "/" + excelFileName +'.xlsx') #loc[:loc.rfind("/")] + excelFileName +'.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 0


books = []

bannedFileTypes = {"jpg", "opf", "db", "tmp", "tmp-journal"}

#for testing use False if there is no limit
stopAfterNumOfBooks = False
def createIndex():
    #create an index so that if there is a need to update the file it does not have to update everything alla agiain

    index = open(os.path.join(excelFileLoc, indexFileName + ".txt" ),"w+")
    index.write(str(books))
    index.close()

def getWebpage(name):
    return(str(requests.get("https://www.goodreads.com/search?utf8=%E2%9C%93&q=" + name + "&search_type=books").content))

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
            f.extend(file)
            if len(f) > stopAfterNumOfBooks and stopAfterNumOfBooks != False: #this is for testing
                return(f)
    return(f)

#methods for getting soting the file name to name and author

def nameAuthor(header):
    name = header[:header.rfind("-")-1]
    author = header[header.rfind("-")+2:]
    return(name, author)

def authorName(header):
    author = header[:header.find("-")]
    book = header[header.rfind("-")+2:]
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
    return(name, author)



def freshUpNameAuthor(name, author):
    name = name.replace(", the", "").replace("_",":")
    author = author.replace(";", ",")
    return(name, author)


def getInfo(header):
    sortingMethods = [nameAuthor, authorName] #this is the differnet methods for getting the name and author so if one fails it's posseble to use a redundant method

    defaultMethod = 1

    fullEntry = header

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
        page = getWebpage(name + " by " + author)

        rating = getRating(page)
        genre = getGenre(page)
        link = getLink(page)

        #its nessecary to have a new method for choosing the method, because this only supports two...
        if rating.find(".") == -1: #choosing the next method
            if defaultMethod == 1:
                defaultMethod = 0
            else:
                defaultMethod = 1
    if getName(page) != "":
        name = getName(page)
        author = getAuthor(page)
    print(defaultMethod)
    return(name, author, format, rating, genre, fullEntry, link)


#this part of the program gets the indexFile
filesInLastBooks = []
pathIndex = excelFileLoc + "/" + indexFileName + ".txt"
if os.path.isfile(pathIndex) and open(pathIndex, "r").read() != "":
    lastBooks = eval(open(pathIndex, "r").read()) #gets the index file
    for i in range(0,len(lastBooks)):
        filesInLastBooks.append(lastBooks[i][5])

i = 0
for file in getAllFilesInDir(loc):
    if (file[file.rfind(".")+1:] in bannedFileTypes) == False:
        if file in filesInLastBooks:
            books.append(lastBooks[filesInLastBooks.index(file)])
        else:
            book = getInfo(file)
            books.append(book)
        if i > 25: #create a backup in the index, so the program can be stopped and resumed and not have to start over
            createIndex()
            i = 0
        else:
            i += 1
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
