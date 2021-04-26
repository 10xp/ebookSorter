# ebookSorter

This program is supposed to list eboks that are in a folder and get the author, the name of the book, rating, genres and link for further information (it scrapes this from goodreads). And the tranfere the data to an excel document so it is easy to sort the books by genre, author and rating. It also creates an index with the raw data, to save and so it is posseble to update the document. The program has gotten quite quick since async functionality was added.

Issues:
 - It has few methods for decifering the file name. Therefore some books don't get rating.
 - The program gets slower after about 70 books
