import xml.etree.ElementTree as ET
import pandas as pd



class api_response:
    def parse_response(self, returnedXML):

        # Parse the XML
        root = ET.fromstring(returnedXML)

        # Namespace mapping
        namespace = {
            'ns': 'http://abr.business.gov.au/ABRXMLSearch/',
            'xsd': 'http://www.w3.org/2001/XMLSchema',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        }

        # Create lists to store extracted data
        data = {
            'ABN': [],
            'EntityName': [],
            'PostCode': [],
            'State': [],
        }

        # Extract data from XML and populate lists...
        for record in root.findall(".//ns:searchResultsRecord", namespace):
            abn_element = record.find(".//ns:ABN/ns:identifierValue", namespace)
            entity_name_element = record.find(".//ns:businessName/ns:organisationName", namespace)
            if entity_name_element is None:
                entity_name_element = record.find(".//ns:mainTradingName/ns:organisationName", namespace)
            if entity_name_element is None:
                entity_name_element = record.find(".//ns:mainName/ns:organisationName", namespace)
            if entity_name_element is None:
                entity_name_element = record.find(".//ns:otherTradingName/ns:organisationName", namespace)   
            state_element = record.find(".//ns:mainBusinessPhysicalAddress/ns:stateCode", namespace)
            postcode_element = record.find(".//ns:mainBusinessPhysicalAddress/ns:postcode", namespace)

            abn = abn_element.text if abn_element is not None else ""
            entity_name = entity_name_element.text if entity_name_element is not None else ""
            state = state_element.text if state_element is not None else ""
            postcode = postcode_element.text if postcode_element is not None else ""

            data['ABN'].append(abn)
            data['EntityName'].append(entity_name)
            data['State'].append(state)
            data['PostCode'].append(postcode)

        # Create DataFrame
        df = pd.DataFrame(data)
        return df 