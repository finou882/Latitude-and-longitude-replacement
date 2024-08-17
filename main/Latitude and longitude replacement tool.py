from openpyxl import load_workbook
from geopy.geocoders import Nominatim
geolocator = Nominatim(user_agent="user-id")
# Load Excel file
excel_path = '../target.xlsx' # Runtime file location in colab (.xlix only)
workbook = load_workbook(filename=excel_path)
# Load sheet
sheet = workbook['Sheet1'] # Name of sheet to which latitude and longitude are to be added
x = 2 # Processing starts from line 2

while x <= 197: # Processing continues until line 54
c1 = sheet['A' + str(x)].value # Get cell value
c2_lat = sheet['B' + str(x)] # Get cell to store latitude
c2_lon = sheet['C' + str(x)] # Get cell to store longitude

if c1: # Make sure c1 is not empty # Start error checking

location = geolocator.geocode(c1)

if location: # Make sure location is not None

c2_lat.value = location.latitude # Save latitude

c2_lon.value = location.longitude # Save longitude

else:

print(f"Could not find location information for '{c1}'.")

else:

print(f"The cell in row {x}, column U is empty.") # End error checking

x += 1 # Increase x by 1

workbook.save(excel_path) # Save changes

workbook.close() # Close workbook
