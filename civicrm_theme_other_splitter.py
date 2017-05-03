import csv
import os

# change these variables as needed
# CSV file exported from CiviCRM - defaults to users desktop
inputfilelocation = os.getenv('USERPROFILE') + "\Desktop\CiviCRM_Activity_Search.csv"
# inputfilelocation = "C:\Users\user.name\Desktop\CiviCRM_Activity_Search.csv"
# Location and file for the results
outputfilelocation = os.getenv('USERPROFILE') + "\Desktop\Themes.csv"

# open an output file
ofile  = open(outputfilelocation, 'wb')
# get set with "standard" excel delimiters
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)

# set an ID counter
counter = 1
# start counting rows
rownum = 0
# open the input file
with open(inputfilelocation, 'rb') as ifile:
  reader = csv.reader(ifile, delimiter=',', quotechar='"')
  # look at each activity (one row from CiviCRM export)
  for row in reader:
    # if this is the header row add some headers to output
    if rownum == 0:
      writer.writerow(["ID","Activity_ID","Theme"])
    # otherwise 
    else:
      # split the "Options" field
      themes = row[1].split(",")
      # for each theme in "Options"
      for theme in themes:
        if theme != '':
          # strip spaces from start/end
          theme = theme.strip()
          # theme for Other provide by "Other" field so skip it
          if theme != 'Other':
            # write the counter, "Activity ID", and theme to output
            writer.writerow([counter,row[0],theme])
            counter += 1
      # if the "Other" field isn't empty
      if row[2] != '':
        # strip spaces from start/end
        s = row[2].strip()
        # write the counter, "Activity ID" and "Other" to output
        writer.writerow([counter,row[0],s])
        counter += 1
    rownum += 1
print rownum," rows processed"
ofile.close()
