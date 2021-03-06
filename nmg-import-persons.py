import pandas
import requests
import urllib
import pycountry

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
# Import Persons
#########################################

# We don't want to import the following names, as they are organizations
company_name_identifiers = ['Limited',
                            'Holding',
                            'Estate',
                            'B.V',
                            'BV',
                            'bv',
                            'Trust',
                            'Plantantions'
                            'School',
                            'Gmbh',
                            'Corporation',
                            'Fund',
                            ]

# Remove rows with empty people
persons_df = df.dropna(how='any', subset=['DIRECTORS/SHAREHOLDERS'])

# Remove names of companies
for company_labels in company_name_identifiers:
    persons_df = persons_df[~persons_df['DIRECTORS/SHAREHOLDERS'].str.contains(
                company_labels)]

# Loop through names of persons
for index, person in persons_df.iterrows():

    pkg = {}
    pkg['@type'] = 'Person'
    pkg['name'] = person['DIRECTORS/SHAREHOLDERS'].replace('\n', ' ')
    pkg['name'] = pkg['name'].strip()
    company = person['Company']
    description = str(person['DESCRIPTION']).title()
    pkg['summary'] = '%s is the %s of %s' % (pkg['name'], description, company)
    if not pandas.isna(person['COUNTRY']):
        try:
            country = pycountry.countries.search_fuzzy(
                    str.strip((person['COUNTRY'])))
            country_token = country[0].numeric
            country_title = country[0].official_name
            pkg['nationalities'] = [{'title': country_title,
                                    'token': country_token}]

        except:
            pass

    # Check if person already exists
    if not search_person_by_name(pkg['name'], session, base_url):
        print(pkg['name'])
        response = session.post('http://localhost:8080/Plone3/persons', json=pkg)

        #import contacts for persons
        if response.status_code == 201 :
            person_url = response.json()['@id']

            if not pandas.isna(person['DIRECTOR/SHAREHOLDER ADDRESS']):
                contact_pkg = {}
                contact_pkg['@type'] = "Contact Detail"
                contact_pkg['label'] = "Director / Shareholder Address"
                contact_pkg['type'] = { 'title': "A postal address", 'token': "address" }
                contact_pkg['value']  = person['DIRECTOR/SHAREHOLDER ADDRESS'].replace('\n','')
                session.post(person_url, json=contact_pkg)

            del response
