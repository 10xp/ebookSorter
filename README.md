# ebookSorter

This program is supposed to list eboks that are in a folder and get the author, the name of the book, rating, genres and link for further information (it scrapes this from goodreads). And then tranfere the data to an excel document so it is easy to sort the books by genre, author and rating. It also creates an index with the raw data, to save and so it is posseble to update the document without finding all books again. The program has gotten quite quick since async functionality was added.

It uses about 6 seconds to get 100 books

Issues:
 - It has few methods for decifering the file name. Therefore some books don't get rating.
 - There are no delays, so goodreads returns empty webpages after a while, 100 books give or take
