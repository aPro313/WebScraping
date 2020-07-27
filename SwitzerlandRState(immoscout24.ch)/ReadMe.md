

#Data extraction from Switzerland real state website immoscout24.ch

#First finding the labels from postal codes with separate script(zipcodes.py). Can be done with below api link for each search
# postcode-5000-aarau?pn=3&r=5
# https://rest-api.immoscout24.ch/v4/en/properties?l=6293&r=5&s=1&t=1

Zipcodes.py file will read the zipcodes from zipcodes.txt file and generate the labels in label.txt file. Then the immoscout.py file will read the label.txt  file and visit the related property cards to get the required information in CSV file. 


