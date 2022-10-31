import googlemaps
from openpyxl import load_workbook
# import pprint

# pp = pprint.PrettyPrinter(indent=4)
wb = load_workbook("./meb.xlsx")
sheet = wb.active
schoolColumn = sheet["C"]

gmaps = googlemaps.Client(key="") # dont push this api key to git services

for count,school in enumerate(schoolColumn):
    # print(school.value)
    try:
        # Geocoding an address
        district_name = sheet["B{}".format(count+1)].value
        full_name = district_name+school.value
        print(count, full_name)
        geocode_result = gmaps.geocode(full_name)
        sheet["D{}".format(count+1)] = (geocode_result[0]["geometry"]["location"]["lat"])
        sheet["E{}".format(count+1)] =  (geocode_result[0]["geometry"]["location"]["lng"])
        # wb.save("./meb.xlsx")
    except IndexError:
        print(IndexError, "[Error at {} in line {}]".format(school.value, count+1))

print("Finished.")

# print(geocode_result[0])
# pp.pprint(geocode_result[0]["geometry"]["location"])

wb.save("./meb.xlsx")
