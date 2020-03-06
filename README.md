# BarcodeDatabase

Database to store Books and Movies in an excel worksheet
works on any platform, but designed to work best on a windows machine
On windows we can open and close excel, and do automatic sorting.

On a non windows we won't be able to sort automatically.

## Books
Books use the calibre program to search by isbn #
Adds in Author, Title, and Series to a worksheet

## Movies
Movies searches moodb, imdb, and upc scavenger for information
Attempts to add Title, Series, and Format to a worksheet
Movies is rather unreliable as there is not a comprehensive site that searches by upc for all movies and is free
