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
NMI_IndustriesArray = ["Retail Trade","Utilities","Arts, Entertainment & Recreation",
                    "Other Services","Health Care & Social Assistance","Accommodation & Food Services",
                    "Finance & Insurance","Real Estate, Rental & Leasing",
                    "Transportation & Warehousing","Mining","Wholesale Trade","Public Administration",
                    "Professional, Scientific & Technical Services","Information","Educational Services",
                    "Management of Companies & Support Services","Construction","Agriculture, Forestry, Fishing & Hunting"]
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

        print("'are:': {}".format(mb3.text.find("are:")))
        print("'report': {}".format(mb3.text.find("report")))
        print("'industries': {}".format(mb3.text.find("industries")))
        print("';': {}".format(mb3.text.find(";")))

        #
        if (mb3.text.find("are:") != -1 and mb3.text.find("report") != -1 and mb3.text.find("industries") != -1 and mb3.text.find(";") != -1 ):
            rankTextArray.append(mb3.text)
            found = True
        if debug == True:
            print("In rankTextArray: {}\n{}".format( found, mb3.text ))
    if grabType(ISM_Page) == "PMI":
        if len(rankTextArray) != 10:
            print("FATAL ERROR: PMI len(rankTextArray) = {}, should be 10! :(".format( len(rankTextArray) ) )
            exit()
        thisSectorsArray = PMI_SectorsArray
        thisIndustriesArray = PMI_IndustriesArray
        
    elif grabType(ISM_Page) == "NMI":
        if len(rankTextArray) != 11:
            print("FATAL ERROR: NMI len(rankTextArray) = {}, should be 11! :(".format( len(rankTextArray) ) )
            exit()
        thisSectorsArray = NMI_SectorsArray
        thisIndustriesArray = NMI_IndustriesArray

    go = input("Press Enter to parse through rankTextArray")

    print("rankTextArray length: {}".format(len(rankTextArray)))
    for rankText in rankTextArray:
        ISM_PositiveRankingList = []
        ISM_NegativeRankingList = []
        ISM_NeutralRankingList = []
        print("\n{}".format(rankText))
        IndustrySplitter = [m.start() for m in re.finditer( "report", str(rankText) )]
        firstRankMin = IndustrySplitter[0]
        secondRankMin = IndustrySplitter[1]

        print("firstRankMin: {}".format(firstRankMin))
        print("secondRankMin: {}".format(secondRankMin))
        positiveRankDictionary = {}
        negativeRankDictionary = {}
        subString1 = rankText[firstRankMin:secondRankMin]
        # print("----")
        # print("subString1: {}".format( subString1 ) )
        for thisIndustry in thisIndustriesArray:
            if subString1.find(thisIndustry) != -1:
                positiveRankDictionary[thisIndustry] = subString1.find( thisIndustry )
                
            subString2 = rankText[secondRankMin:]
            if subString2.find(thisIndustry) != -1:
                negativeRankDictionary[thisIndustry] = subString2.find( thisIndustry )
                
        sortedPositiveRankList = sorted( positiveRankDictionary.items(), key=lambda kv: kv[1])
        sortedNegativeRankList = sorted( negativeRankDictionary.items(), key=lambda kv: kv[1])
        print("sortedPositiveRankList: {}".format( sortedPositiveRankList ) )
        print("sortedNegativeRankList: {}".format( sortedNegativeRankList ) )

        debugRankList = False

        # ADD NEUTRAL INDUSTRIES TO ENTIRE ARRAY
        for neutralIndustry in thisIndustriesArray:
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

    print("len(rankTextArray): {}".format( len(rankTextArray) ) )
    finalRankDictionaryArray = [ dict() for x in range(len(rankTextArray) ) ] 
    
    for dictIndex, dictionary in enumerate(finalRankDictionaryArray):
        print("dictIndex: {}".format(dictIndex))
        for industryIndex, industry in enumerate(ISM_PositiveRankingListArray[dictIndex]):
            dictionary[industry] = industryIndex + 1

        neutralOffset = len(ISM_PositiveRankingListArray[dictIndex])
        for industryIndex2, industry2 in enumerate(ISM_NeutralRankingListArray[dictIndex]):
            dictionary[industry2] = neutralOffset
        
        negativeOffset = neutralOffset + len(ISM_NeutralRankingListArray[dictIndex])
        ISM_NegativeRankingListArray[dictIndex].reverse()
        for industryIndex3, industry3 in enumerate(ISM_NegativeRankingListArray[dictIndex]):
            dictionary[industry3] = negativeOffset + industryIndex3 + 1
        try:
            go = input("Press enter for dictionary {}".format(thisSectorsArray[dictIndex]))
            print(dictionary)
            print("")
        except IndexError:
            break
            
            
        
    return finalRankDictionaryArray

if __name__ == "__main__":
    grabISMcomments("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")
    # grabISMrankings("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")