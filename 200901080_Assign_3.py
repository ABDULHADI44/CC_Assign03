import xml.etree.ElementTree as ET
import openpyxl

# parse the XML file
tree = ET.parse("C:/Users/Dell/Downloads/compiler.xml")

# get the root element of the XML document
root = tree.getroot()

# create a list to store the extracted data
data = []

# iterate over the book elements in the XML file
for book in root.findall("book"):
    # create a dictionary to store the data for each book
    book_data = {}
    # extract the data for each book
    book_data["Book_Id"] = book.get("id")
    book_data["Author_Name"] = book.find("author").text
    book_data["Title"] = book.find("title").text
    book_data["Genre"] = book.find("genre").text
    book_data["Price"] = book.find("price").text
    book_data["Publish_date"] = book.find("publish_date").text
    book_data["Description"] = book.find("description").text
    # add the book data to the list
    data.append(book_data)


# create a new Excel file
wb = openpyxl.Workbook()

# create a sheet in the Excel file
sheet = wb.active

# add the column headings to the sheet
sheet.append(["Book_Id", "Author_Name", "Title", "Genre", "Price", "Publish_date", "Description"])

# add the extracted data to the sheet
for book in data:
    sheet.append([book["Book_Id"], book["Author_Name"], book["Title"], book["Genre"], book["Price"], book["Publish_date"], book["Description"]])

# save the Excel file
wb.save("200901080_Assignment_03.xlsx")
