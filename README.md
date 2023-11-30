# KML-From-Excel
Generate KML file from Excel file of addresses

Excel (xlsx format) should have the following columns
Call, Address, City, State, ZipCode, Opt-In

Call: Ham radio call sign used as label on KML map point
Address: number and street
City: name of city
State: state where city is located
ZipCode: 5 digit U.S. zip code
Opt-In: set to 'Y' or 'N'

For Opt-In = 'Y', those KML points will use the full precision of lat & long from Google Maps
For Opt-In = 'N', those KML points will have their precision reduced to two decimal places

Make sure to save your Excel file, with the columns above, with the name addresses.xlsx so
the program can find it