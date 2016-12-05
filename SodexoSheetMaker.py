import xlrd, os, csv, openpyxl, datetime
from shutil import copyfile

# Counter for which row in sodexo sheet to put this
count = 2

# Link to the good drive folder
facebookTwitterImages = "https://drive.google.com/open?id=0B8Ai3j-fa-oqZjdsd19YVDZseHM"

def main():
    global count
    
    # Make a duplicate of the template 
    copyfile('files/SodexoSheetTemplate.xlsx', 'files/CompletedSheet.xlsx')

    # Open the template and open the first sheet
    sodexoSheet = openpyxl.load_workbook('files/CompletedSheet.xlsx')
    sodexoSheet_sheet = sodexoSheet.worksheets[0]

    # Loop through the other files in the directory and do certain tasks with them
    for file in os.listdir('files'):
        file = 'files/' + file
        if (file == 'CompletedSheet.xlsx' or file == 'SodexoSheetTemplate.xlsx'):
            continue
            
        if('tweet' in file.lower()):
            twitter(file, sodexoSheet_sheet)
        elif("facebook" in file.lower()):
            facebook(file, sodexoSheet_sheet)
            
    # Save the sheet when done
    sodexoSheet.save('files/CompletedSheet.xlsx')
        
def twitter(file, sodexoSheet_sheet):
    global count
    # Open the twitter export
    with open(file, 'rb') as f:
        twitterReader = csv.reader(f)

        # Iterate through all of the rows of the sheet and get all of the values we want
        for row in twitterReader:            
            if(count == 2):
                count += 1
                continue
            
            # Get all of the values needed
            caption = row[2]
            dateTime = row[3]
            dateTimeList = dateTime.split(' ')
            date = dateTimeList[0]
            time = dateTimeList[len(dateTimeList)-2]
            impressions = row[4]
            
            # Check if the caption is a reply and if it is, don't include it
            if(caption[0] == '@'):
                continue

            sodexoSheet_sheet.cell(row=count , column=1).value = date
            sodexoSheet_sheet.cell(row=count , column=2).value = time
            sodexoSheet_sheet.cell(row=count , column=3).value = caption
            sodexoSheet_sheet.cell(row=count , column=4).value = impressions
            sodexoSheet_sheet.cell(row=count, column=6).value = facebookTwitterImages
            sodexoSheet_sheet.cell(row=count, column=7).value = 'TWITTER'
            count+=1

def facebook(file, sodexoSheet_sheet):
    global count
    
    # Open the workbook
    facebookReader = xlrd.open_workbook(file)
    
    # Open the sheet
    facebookReader_sheet = facebookReader.sheet_by_index(0)
    
    row = 2
    while(row < facebookReader_sheet.nrows):
        if(facebookReader_sheet.cell(row, 2).value == ''):
            return
        
        caption = facebookReader_sheet.cell(row,2).value
        impressions = facebookReader_sheet.cell(row,11).value
        dateTime = facebookReader_sheet.cell(row,6).value
        dateTimeAsDateTime = datetime.datetime(*xlrd.xldate_as_tuple(dateTime, facebookReader.datemode))
        date = str(dateTimeAsDateTime.date())
        timeFull = str(dateTimeAsDateTime.time())
        timeList = timeFull.split(':')
        time = timeList[0] + ":" + timeList[1]
        #print date

        sodexoSheet_sheet.cell(row=count , column=1).value = date
        sodexoSheet_sheet.cell(row=count, column=2).value = time
        sodexoSheet_sheet.cell(row=count , column=3).value = caption
        sodexoSheet_sheet.cell(row=count , column=4).value = impressions
        sodexoSheet_sheet.cell(row=count, column=6).value = facebookTwitterImages
        sodexoSheet_sheet.cell(row=count, column=7).value = 'FACEBOOK'
        count+=1
        row+=1
            
main()