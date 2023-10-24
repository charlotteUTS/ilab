import urllib.request as req



"""
class abn:
    def __init__(self, keyword, postcode):
        self.name = keyword
        self.postcode = postcode
        self.base_url = 'https://abr.business.gov.au/ABRXMLSearch/AbrXmlSearch.asmx'
        self.legalName = ''		
        self.tradingName = ''	
        self.NSW = 'Y'			
        self.SA = 'N'				
        self.ACT = 'N'			
        self.VIC = 'N'			
        self.WA = 'N'				
        self.NT = 'N'				
        self.QLD = 'N'			
        self.TAS = 'N'			
        self.authenticationGuid = '3ecee520-acf0-4b47-80f2-25ee15f00bc9'
        
        
    def open(self):
        url = 'https://abr.business.gov.au/abrxmlsearchRPC/AbrXmlSearch.asmx/ABRSearchByNameSimpleProtocol?name=' + self.name + '&postcode=' + self.postcode + '&legalName=' + self.legalName + '&tradingName=' + self.tradingName + '&NSW=' + self.NSW + '&SA=' + self.SA + '&ACT=' + self.ACT + '&VIC=' +  self.VIC + '&WA=' + self.WA + '&NT=' + self.NT + '&QLD=' + self.QLD + '&TAS=' + self.TAS + '&authenticationGuid=' + self.authenticationGuid
        response = req.urlopen(url)
        return response

"""

class abn:
    def __init__(self, keyword, postcode):
        self.name = keyword.replace(" ", "+")
        print(keyword)
        self.postcode = postcode
        self.base_url = 'https://abr.business.gov.au/ABRXMLSearch/AbrXmlSearch.asmx'
        self.legalName = ''		
        self.tradingName = ''	
        self.NSW = 'Y'					
        self.authenticationGuid = '3ecee520-acf0-4b47-80f2-25ee15f00bc9'
        self.activeABNsOnly = 'Y' 
    


    def open(self):
        url = 'https://abr.business.gov.au/abrxmlsearch/AbrXmlSearch.asmx/ABRSearchByNameAdvancedSimpleProtocol2017?name=' + self.name + '&postcode=' + self.postcode +'&legalName=&tradingName=&businessName=&activeABNsOnly=' + self.activeABNsOnly + '&NSW=Y&SA=N&ACT=N&VIC=N&WA=N&NT=N&QLD=N&TAS=N&authenticationGuid=' + self.authenticationGuid + '&searchWidth=&minimumScore=&maxSearchResults=1000000'

        response = req.urlopen(url)
        return response


    






