


#This program will find the names, athors, rating for the ebboks in a folder

import os
import requests
import xlsxwriter

#input
loc = "//OMV2M/Publicmappe/bøker/Sci-Fi.and.Fantasy.Ebook.Collection"
excelFileName = "Sci-Fi.and.Fantasy.Ebook.Collection"

#excel-variables
workbook = xlsxwriter.Workbook(loc[:loc.rfind("/")+1] + excelFileName +'.xlsx') #loc[:loc.rfind("/")] + excelFileName +'.xlsx')
worksheet = workbook.add_worksheet()

row = 1
col = 0


books = []


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
        name = name + "(" + series + ")"

    name.replace(" -",":")
    return(name, author)



def freshUpNameAuthor(name, author):
    name = name.replace(", the", "").replace("_",":")
    author = author.replace(";", ",")
    return(name, author)



def getInfo(header, lastMethod):
    sortingMethods = [nameAuthor, authorName] #this is the differnet methods for getting the name and author so if one fails it's posseble to use a redundan method

    format = header[header.rfind(".")+1:]
    header = header[:header.rfind(".")]
    rating = ""

    i = 0
    while rating.find(".") == -1 and i < len(sortingMethods): #making sure that its not empty.
        i+=1
        temp = sortingMethods[lastMethod](header)

        name = temp[0]
        author = temp[1]

        #this uses the website goodreads to get the data
        page = getWebpage(name + "-" + author)

        rating = getRating(page)
        genre = getGenre(page)
        link = getLink(page)

        #its nessecary to have a new method for choosing the method, because this only supports two...
        if rating.find(".") == -1: #choosing the next method
            if lastMethod == 1:
                lastMethod = 0
            else:
                lastMethod = 1
    if getName(page) != "":
        name = getName(page)
        author = getAuthor(page)
    return(name, author, format, rating, genre, link)

with os.scandir(loc) as entries:
    lastMethod = 0
    i = 0 #for testing
    for entry in entries:
        book = getInfo(entry.name, lastMethod)
        books.append(book)


        #for testing
        i+= 1
        if i > 100:
            pass
            #break


#create a xcel document

worksheet.add_table('A1:E1000', {'data': books, "columns": [{'header':"Book"}, {'header':"Author(s)"}, {'header':"Format"}, {'header':"Rating"}, {'header':"Genre"}]})
for i in range(0, len(books)): #makes the name a link to get more information
    worksheet.write_url(row, 0, books[i][5], string=books[i][0])
    row+=1
    pass
workbook.close()
