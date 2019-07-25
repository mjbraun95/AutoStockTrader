from bs4 import BeautifulSoup
import operator
import re

PMI_IndustriesArray = ["Machinery","Computer & Electronic Products","Paper Products",
                    "Apparel, Leather & Allied Products","Printing & Related Support Activities",
                    "Primary Metals","Nonmetallic Mineral Products","Petroleum & Coal Products",
                    "Plastics & Rubber Products","Miscellaneous Manufacturing",
                    "Food, Beverage & Tobacco Products","Furniture & Related Products",
                    "Transportation Equipment","Chemical Products","Fabricated Metal Products",
                    "Electrical Equipment, Appliances & Components","Textile Mills","Wood Products"]
NMI_IndustriesArray = ["Retail Trade","Utilities","Arts, Entertainment Recreation",
                    "Other Services","Healthcare and Social Assistance","Food and Accomodations",
                    "Finance and Insurance","Real Estate, Renting and Leasing",
                    "Transport and Warehouse","Mining","Wholesale","Public Admin",
                    "Professional, Science and Technology Services","Information","Education",
                    "Management","Construction","Agriculture, Forest, Fishing and Hunting"]
PMI_SectorsArray = ["NEW ORDERS","PRODUCTION","EMPLOYMENT","SUPPLIER DELIVERIES","INVENTORIES",
                    "CUSTOMER INVENTORIES","PRICES","BACKLOG OF ORDERS","EXPORTS","IMPORTS"]
NMI_SectorsArray = ["ISM NON-MANUFACTURING","BUSINESS ACTIVITY","NEW ORDERS","EMPLOYMENT","SUPPLIER DELIVERIES","INVENTORIES",
                    "PRICES","BACKLOG OF ORDERS","EXPORTS","IMPORTS"]

SectorsArray = [""]

debug = True

def grabType(ISM_Page):
    ISM_Soup = BeautifulSoup(ISM_Page.content, 'html.parser')
    tcts_Array = ISM_Soup.findAll("h4",{"class":"text-center text-strong"})
    print("len(tcts_Array): {}".format(len(tcts_Array)))
    print(tcts_Array[0].text)
    if tcts_Array[0].text.find("PMI") != -1:
        print("PMI page detected!")
        Type = "PMI"
    elif tcts_Array[0].text.find("NMI") != -1:
        print("NMI page detected!")
        Type = "NMI"
    else:
        print("ERROR: NOT PMI OR NMI. ASSUMING PMI")
        Type = "PMI"
    return Type

def grabISMcomments(ISM_Page):
    ISM_Soup = BeautifulSoup(ISM_Page.content, 'html.parser')
    print("Grabbing ISM Comments...")
    lgi_Array = ISM_Soup.findAll("li",{"class":"list-group-item"})
    ISM_CommentList = []
    x = 6
    while x < len(lgi_Array):
        ISM_CommentList.append(lgi_Array[x-6])
        x += 1

    # print("Data array created from {} to {}!".format(ISM_CommentList[0][0],ISM_CommentList[-1][0]))
    print("Number of notes: {}".format(len(ISM_CommentList)))
    return ISM_CommentList
    # return 0

def grabISMrankings(ISM_Page):
    ISM_Soup = BeautifulSoup(ISM_Page.content, 'html.parser')
    print("Grabbing ISM Rankings...")
    ISM_RankingList = []
    ISM_PositiveRankingList = []
    ISM_NegativeRankingList = []
    ISM_NeutralRankingList = []
    ISM_PositiveRankingListArray = []
    ISM_NegativeRankingListArray = []
    ISM_NeutralRankingListArray = []
    rankTextArray = []
    mb3_Array = ISM_Soup.findAll("p",{"class":"mb-3"})
    if debug == True:
        print("\n\nlen(mb3_Array): {}".format( len(mb3_Array) ))
        
    for mb3_i, mb3 in enumerate(mb3_Array):
        found = False
        # print("{}.find({}): {}".format( mb3.text, keyword, mb3.text.find(keyword)))
        if (mb3.text.find("order") != -1 and mb3.text.find("report") != -1 and mb3.text.find("industries") != -1 and mb3.text.find(";") != -1 ):
            rankTextArray.append(mb3.text)
            found = True
        if debug == True:
            print("\nIn rankTextArray: {}\n{}".format( found, mb3.text ))
    if grabType(ISM_Page) == "PMI":
        if len(rankTextArray) != 10:
            print("FATAL ERROR: PMI len(rankTextArray) = {}, should be 10! :(".format( len(rankTextArray) ) )
            exit()
        thisSectorsArray = PMI_SectorsArray
        
    elif grabType(ISM_Page) == "NMI":
        if len(rankTextArray) != 11:
            print("FATAL ERROR: NMI len(rankTextArray) = {}, should be 11! :(".format( len(rankTextArray) ) )
            exit()
        thisSectorsArray = NMI_SectorsArray

    print("rankTextArray length: {}".format(len(rankTextArray)))
    for rankText in rankTextArray:
        ISM_PositiveRankingList = []
        ISM_NegativeRankingList = []
        ISM_NeutralRankingList = []
        print("\n{}".format(rankText))
        colons = [m.start() for m in re.finditer( ":", str(rankText) )]
        firstRankMin = colons[0]
        secondRankMin = colons[1]

        # thirdRankOrder = rankText[secondRankMin+1:].find(":")
        print("firstRankMin: {}".format(firstRankMin))
        print("secondRankMin: {}".format(secondRankMin))
        # print("thirdRankOrder: {}".format(thirdRankOrder))
        positiveRankDictionary = {}
        negativeRankDictionary = {}
        # ISM_NeutralRankingList = PMI_IndustriesArray
        subString1 = rankText[firstRankMin:secondRankMin]
        # print("----")
        # print("subString1: {}".format( subString1 ) )
        for PMI_Industry in PMI_IndustriesArray:
            if subString1.find(PMI_Industry) != -1:
                positiveRankDictionary[PMI_Industry] = subString1.find( PMI_Industry )
                
            subString2 = rankText[secondRankMin:]
            if subString2.find(PMI_Industry) != -1:
                negativeRankDictionary[PMI_Industry] = subString2.find( PMI_Industry )
                
        sortedPositiveRankList = sorted( positiveRankDictionary.items(), key=lambda kv: kv[1])
        sortedNegativeRankList = sorted( negativeRankDictionary.items(), key=lambda kv: kv[1])
        print("sortedPositiveRankList: {}".format( sortedPositiveRankList ) )
        print("sortedNegativeRankList: {}".format( sortedNegativeRankList ) )

        debugRankList = False

        # ADD NEUTRAL INDUSTRIES TO ENTIRE ARRAY
        for neutralIndustry in PMI_IndustriesArray:
            ISM_NeutralRankingList.append( neutralIndustry )

        # POSITIVE INDUSTRIES
        for index, positiveIndustry in enumerate( sortedPositiveRankList ):
            
            quotes = [m.start() for m in re.finditer( "'", str(positiveIndustry) )]
            posStrL = quotes[0]
            posStrR = quotes[1]
            posStr = str(positiveIndustry)[posStrL+1:posStrR]

            if debugRankList == True:
                print("posStr: {}".format( posStr ) )
            ISM_PositiveRankingList.append( posStr )

            # Removes positives from neutral
            ISM_NeutralRankingList.remove( posStr )

        # NEGATIVE INDUSTRIES
        for index, negativeIndustry in enumerate( sortedNegativeRankList ):
            
            quotes = [m.start() for m in re.finditer( "'", str(negativeIndustry) )]
            negStrL = quotes[0]
            negStrR = quotes[1]
            negStr = str(negativeIndustry)[negStrL+1:negStrR]

            if debugRankList == True:
                print("negStr: {}".format( negStr ) )
            ISM_NegativeRankingList.append( negStr )

            # Removes negatives from neutral
            ISM_NeutralRankingList.remove( negStr )

        ISM_PositiveRankingListArray.append( ISM_PositiveRankingList )
        ISM_NegativeRankingListArray.append( ISM_NegativeRankingList )
        ISM_NeutralRankingListArray.append( ISM_NeutralRankingList )

        # for i in range(9):
        print("----")
        print("ISM_PositiveRankingListArray[-1]: {}".format( ISM_PositiveRankingListArray[-1] ) )
        print("ISM_NegativeRankingListArray[-1]: {}".format( ISM_NegativeRankingListArray[-1] ) )
        print(" ISM_NeutralRankingListArray[-1]: {}".format( ISM_NeutralRankingListArray[-1] ) )
        print("----")
    finalRankDictionaryArray = [ dict() for x in range(len(rankTextArray) ) ] 
    
    for dictIndex, dictionary in enumerate(finalRankDictionaryArray):
        for industryIndex, industry in enumerate(ISM_PositiveRankingListArray[dictIndex]):
            dictionary[industry] = industryIndex + 1

        neutralOffset = len(ISM_PositiveRankingListArray[dictIndex])
        for industryIndex2, industry2 in enumerate(ISM_NeutralRankingListArray[dictIndex]):
            dictionary[industry2] = neutralOffset
        
        negativeOffset = neutralOffset + len(ISM_NeutralRankingListArray[dictIndex])
        ISM_NegativeRankingListArray[dictIndex].reverse()
        for industryIndex3, industry3 in enumerate(ISM_NegativeRankingListArray[dictIndex]):
            dictionary[industry3] = negativeOffset + industryIndex3 + 1

        go = input("Press enter for dictionary {}".format(thisSectorsArray[dictIndex]))
        print(dictionary)
        print("")
            
            
        
    return finalRankDictionaryArray

if __name__ == "__main__":
    # grabISMcomments("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")
    grabISMrankings("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")