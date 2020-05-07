import xlsxwriter       #xlswriter is a python library for creating spreadsheets
from PIL import Image   #Pillow is a python library for image processing. It contains PIL. !DO NOT INSTALL THE PIL LIBRARY
from operator import itemgetter     #to extract the elements of tuples from a list of tuples into different lists depending on their position in the tuple 

img = Image.open('your_image_file_name.image_file_extension')      #opens the given picture file
# img = img.convert('RGB')      #optional to remove alpha values from RGBA
size = img.size                 #returns the picture size or dimensions as a tuple
row, col = size[0], size[1]     #stores the first element of given tuple in row and second element in col

row_limit, col_limit = 100, 100     #set the pixel limit for picture. If the picture size exceeds these limits it gets resized 

resized = False     #to check whether the picture got resized
if row > row_limit or col > col_limit:
    img = img.resize((row_limit, col_limit))    #resizes the picture if it crosses the maximum pixels limit
    resized = True  #tags the picture as resized

if not resized:     #checks whether the picture is resized or not
    row = 3*row     #since 3 vertical cells of a spreadsheeet make one pixel we have to multiply number of rows with three
else:
    row = 3*row_limit       #if picture is resized we do the same thing but with the new dimensions
    col = col_limit

data = list(img.getdata())      #returns a list of tuples where each tuple contains RGBA values of one pixel. Pixels get read from Top left to right 

redlist = list(map(itemgetter(0), data))        #returns a list of R values of each pixel. In other words the first element of each tuple in the data list
greenlist = list(map(itemgetter(1), data))      #returns a list of G values of each pixel. In other words the second element of each tuple in the data list
bluelist = list(map(itemgetter(2), data))       #returns a list of B values of rach pixel. In other wprds the third element of each tuple in the data list

workbook = xlsxwriter.Workbook("your_excel_file_name.xlsx")        #creates a new excel file

worksheet = workbook.add_worksheet()        #adds a new worksheet to the excel file

redx, greenx, bluex = 0, 1, 2       #to track the row numbers in order to format them with correct color. Red rows are [0, 3, 6,...], Green rows are [1, 4, 7,...], Blue rows are [2, 5, 8,...]

iteration_num = row//3      #since we are formating 3 rows at a time, we take the original pixel count or original rows count as range for iteration 
for _ in range(iteration_num):
    #this format the red rows. For each row, the formating starts from the right and ends at the nth cell where n is the number of columns
    worksheet.conditional_format(redx, 0, redx, col, {
        'type': '2_color_scale',
        'min_type': 'num',
        'max_type': 'num',
        'min_value': 0,
        'max_value': 255,
        'min_color': '#000000',
        'max_color': '#FF0000',
    })
    
    #this formats the green rows.
    worksheet.conditional_format(greenx, 0, greenx, col, {
        'type': '2_color_scale',
        'min_type': 'num',
        'max_type': 'num',
        'min_value': 0,
        'max_value': 255,
        'min_color': '#000000',
        'max_color': '#00FF00',
    })
    
    #this formats the blue rows
    worksheet.conditional_format(bluex, 0, bluex, col, {
        'type': '2_color_scale',
        'min_type': 'num',
        'max_type': 'num',
        'min_value': 0,
        'max_value': 255,
        'min_color': '#000000',
        'max_color': '#0000FF',
    })

    #after each iteration we increase the respective rows by 3 to get to the next set of three rows or say next pixel below the current one
    redx += 3
    greenx += 3
    bluex +=3

cellx, celly = 0, 0     #to track in which cell we are at any given time. cellx corresponds to the row number and celly corresponds to the column number of a particular cell
count = 0           #to track what row type we are in. 0 corresponds to red, 1 corresponds to green, and 2 corresponds to blue
r, g, b = 0, 0, 0       #to grab RGB values from their respective lists. Each time we grab a value from a list the variable gets increased by one depemding on the fact that from which list we grabbed the value from 

for _ in range(row): #each iteration formats a single row
    for _ in range(col): #each iteration formats a single cell
        if count == 0:  #checks whether the row is to be formatted red or not
            worksheet.write(cellx, celly, redlist[r])      #writes the rth R vlaue from the redlist list on the cell whose address is located using cellx and celly 
            r += 1      #r is increased by one so that next time the next R value from the redlist gets written in the next cell
        elif count == 1:    #checks whether the row is to be formated green or not
            worksheet.write(cellx, celly, greenlist[g])     #writes the gth G vlaue from the greenlist list on the cell whose address is located using cellx and celly
            g += 1      #g is increased by one so that next time the next G value from the redlist gets written in the next cell
        elif count == 2:    #checks whether the row is to be formatted blue or not
            worksheet.write(cellx, celly, bluelist[b])      #writes the bth B vlaue from the bluelist list on the cell whose address is located using cellx and celly
            b += 1      #b is increased by one so that next time the next B value from the redlist gets written in the next cell
        celly += 1
    if count == 2:      #to reset the count back to 0 on formating the third row of a row set(equals to one vertical pixel)
        count = 0
    else:               #if the row set is not formatted compelety count get increased by one. A row set is said to be completed when the blue row which is the third one in a row set get formatted
        count += 1
    celly = 0       #to reset the column number back to zero so that we can start formatng from right again
    cellx += 1      #to move on to the next row. Doing this three times completes the formating of one row set or one vertical pixel

workbook.close()    #closes the excel file