import pandas
import requests
import urllib
import pycountry
import datetime

# Functions


def search_person_by_name(name, session, base_url):
    # Searches for exact match of name (title)
    # If true, returns dict

    # Workaround for bug
    # https://github.com/plone/plone.restapi/issues/998
    safename = name.replace('(', '')
    safename = safename.replace(')', '')

    url = base_url + '/@search' + \
                     '?Title=' + urllib.parse.quote_plus(safename) + \
                     '&portal_type%3Alist=Person'

    response = session.get(url)

    try:
        if not response.json()['items']:
            return None

        if response.json()['items'][0]['title'] == name:
            return response.json()['items'][0]

        else:
            return None

    except:
        pass
        
def search_organization_by_name(name, session, base_url):
    # Searches for best text match of name (SearchableText)
    # If true, returns dict of first result

    # Workaround for bug
    # https://github.com/plone/plone.restapi/issues/998
    safename = name.replace('(', '')
    safename = safename.replace(')', '')

    url = base_url + '/@search' + \
                     '?Title=' + urllib.parse.quote_plus(safename) + \
                     '&portal_type%3Alist=Organization'

    response = session.get(url)

    try:
        if not response.json()['items']:
            return None

        if response.json()['items'][0]['title'] == name:
            return response.json()['items'][0]

        else:
            return None

    except:
        pass


# Setup authenticated session
# https://plonerestapi.readthedocs.io/en/latest/authentication.html

base_url = "http://localhost:8080/Plone3"

session = requests.Session()
session.auth = ('', '')
session.headers.update({'Accept': 'application/json'})

# Import Dataframe
# This script uses Pandas library to handle import of spreadsheet data,
# including clean up operations
# https://pandas.pydata.org/pandas-docs/stable/getting_started/index.html

df = pandas.read_excel('source/kenya-companies.xlsx')

# replace UK in countries with United Kingdom
df['COUNTRY'].replace({'UK': "United Kingdom"}, inplace=True)
#########################################
# Import Organizations
#########################################

# Remove rows with empty companies 
orgs_df = df.dropna(how='any', subset=['Company'])

# Loop through names of orgs

for index, org in orgs_df.iterrows():

    pkg = {}
    pkg['@type'] = 'Organization'
    pkg['name'] = org['Company'].replace('\n', ' ')
    pkg['name'] = pkg['name'].strip()

    if type(org['DATE OF REGISTRATION']) == datetime.date:
        pkg['founding_date'] = org['DATE OF REGISTRATION'].isoformat()

    # description 
    # company is only for shareholders at this time

    pkg['description'] = u''

    # Check if organization already exists
    if not search_organization_by_name(pkg['name'], session, base_url):
        print(pkg['name'])
        response = session.post('http://localhost:8080/Plone3/orgs', json=pkg)

        if response.status_code == 201:
            org_url = response.json()['@id']

            # import addressess

            # registered offices
            # Needs delimiters
            # if not pandas.isna(org['REGISTERED OFFICE']):
            #   contact_pkg = {}
            #    contact_pkg['@type'] = "Contact Detail"
            #    contact_pkg['label'] = "Registered Office"
            #    contact_pkg['type'] = {'title': "A postal address", 'token': "address"}
            #    contact_pkg['value']  = org['REGISTERED OFFICE'].replace('\n', '')
            #    session.post(org_url, json=contact_pkg)

            # copmany postal address

            if not pandas.isna(org['COMPANY POSTAL ADDRESS']):
                contact_pkg = {}
                contact_pkg['@type'] = "Contact Detail"
                contact_pkg['label'] = "Company Postal Address"
                contact_pkg['type'] = { 'title': "A postal address", 'token': "address"}
                contact_pkg['value']  = org['COMPANY POSTAL ADDRESS'].replace('\n','')
                session.post(org_url, json=contact_pkg)

            # import identifier

            if not pandas.isna(org['COMPANY NUMBER']):
                    identifier_pkg = {}
                    identifier_pkg["@type"] = "Identifier"
                    identifier_pkg['scheme'] = "Company Number"
                    identifier_pkg['identifier'] = org['COMPANY NUMBER']
                    session.post(org_url, json=identifier_pkg)
            del response




