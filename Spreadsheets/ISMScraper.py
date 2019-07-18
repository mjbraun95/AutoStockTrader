import requests
from bs4 import BeautifulSoup

debug = True

def grabISMnotes():
    # URL = input("Enter URL Here: ")
    ISM_Page = requests.get("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")#https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")
    print("status code: {}".format(ISM_Page.status_code))

    ISM_Soup = BeautifulSoup(ISM_Page.content, 'html.parser')
    notelist = ISM_Soup.findAll("li",{"class":"list-group-item"})
    ISMnotelist = []
    x = 6
    while x < len(notelist):
        ISMnotelist.append(notelist[x-6])
        x += 1
    # for note in notelist:
        # if note.text ==
        #     noteText = note.text
        #     ISMnotelist.append(noteText)

        # if debug == True:
        #     print("")
        #     print("noteText: {}".format(noteText))
        #     print(" ISMnotelist length: {}".format(len(ISMnotelist)))
        #     print()
    
    # print("Data array created from {} to {}!".format(ISMnotelist[0][0],ISMnotelist[-1][0]))
    print("Number of notes: {}".format(len(ISMnotelist)))
    return ISMnotelist
    # return 0

if __name__ == "__main__":
    grabISMnotes()