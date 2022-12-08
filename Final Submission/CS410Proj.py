
#import the required packages
#xlwt is for generating excel file, requests is for send request and reiceve the returned data, esaygui is for input gui 
import xlwt
import requests
import easygui as gui



# Define a method with offset, limit, location, row
def run(offset, limit, location, row):

    page = 0
    # Use loop to go through all pages
    while True:
        #ask the required parameters
        json_data = {
            'query': '\n\nquery ConsumerSearchMainQuery($query: HomeSearchCriteria!, $limit: Int, $offset: Int, $sort: [SearchAPISort], $sort_type: SearchSortType, $client_data: JSON, $bucket: SearchAPIBucket)\n{\n  home_search: home_search(query: $query,\n    sort: $sort,\n    limit: $limit,\n    offset: $offset,\n    sort_type: $sort_type,\n    client_data: $client_data,\n    bucket: $bucket,\n  ){\n    count\n    total\n    results {\n      property_id\n      list_price\n      primary\n      primary_photo (https: true){\n        href\n      }\n      source {\n        id\n        agents{\n          office_name\n        }\n        type\n        spec_id\n        plan_id\n      }\n      community {\n        property_id\n        description {\n          name\n        }\n        advertisers{\n          office{\n            hours\n            phones {\n              type\n              number\n            }\n          }\n          builder {\n            fulfillment_id\n          }\n        }\n      }\n      products {\n        brand_name\n        products\n      }\n      listing_id\n      matterport\n      virtual_tours{\n        href\n        type\n      }\n      status\n      permalink\n      price_reduced_amount\n      other_listings{rdc {\n      listing_id\n      status\n      listing_key\n      primary\n    }}\n      description{\n        beds\n        baths\n        baths_full\n        baths_half\n        baths_1qtr\n        baths_3qtr\n        garage\n        stories\n        type\n        sub_type\n        lot_sqft\n        sqft\n        year_built\n        sold_price\n        sold_date\n        name\n      }\n      location{\n        street_view_url\n        address{\n          line\n          postal_code\n          state\n          state_code\n          city\n          coordinate {\n            lat\n            lon\n          }\n        }\n        county {\n          name\n          fips_code\n        }\n      }\n      tax_record {\n        public_record_id\n      }\n      lead_attributes {\n        show_contact_an_agent\n        opcity_lead_attributes {\n          cashback_enabled\n          flip_the_market_enabled\n        }\n        lead_type\n        ready_connect_mortgage {\n          show_contact_a_lender\n          show_veterans_united\n        }\n      }\n      open_houses {\n        start_date\n        end_date\n        description\n        methods\n        time_zone\n        dst\n      }\n      flags{\n        is_coming_soon\n        is_pending\n        is_foreclosure\n        is_contingent\n        is_new_construction\n        is_new_listing (days: 14)\n        is_price_reduced (days: 30)\n        is_plan\n        is_subdivision\n      }\n      list_date\n      last_update_date\n      coming_soon_date\n      photos(limit: 2, https: true){\n        href\n      }\n      tags\n      branding {\n        type\n        photo\n        name\n      }\n    }\n  }\n}',
            'variables': {
                'query': {
                    'status': [
                        'for_sale',
                        'ready_to_build',
                    ],
                    'primary': True,
                    'search_location': {
                        'location': location,#City
                    },
                },
                'client_data': {
                    'device_data': {
                        'device_type': 'web',
                    },
                    'user_data': {
                        'last_view_timestamp': -1,
                    },
                },
                'limit': limit,#The number of rows returned per page
                'offset': offset,
                'zohoQuery': {
                    'silo': 'search_result_page',
                    'location': location,#City
                    'property_status': 'for_sale',
                    'filters': {},
                    'page_index': '1',
                },
                'sort_type': 'relevant',
                'geoSupportedSlug': location,#City
                'bucket': {
                    'sort': 'modelF',
                },
                'resetMap': '2022-11-24T03:56:12.438Z0.36206071581266985',
                'by_prop_type': [
                    'home',
                ],
            },
            'operationName': 'ConsumerSearchMainQuery',
            'callfrom': 'SRP',
            'nrQueryType': 'MAIN_SRP',
            'visitor_id': '0e04d5de-20ea-4926-bcf6-49428c17dc97',
            'isClient': True,
            'seoPayload': {
                'asPath': '/realestateandhomes-search/San-Diego_CA/pg-58',
                'pageType': {
                    'silo': 'search_result_page',
                    'status': 'for_sale',
                },
                'county_needed_for_uniq': False,
            },
        }

        # Initialize a request to get one page of data
        response = requests.post('https://www.realtor.com/api/v1/hulk_main_srp', params=params, cookies=cookies,
                                 headers=headers, json=json_data)
        # Turn returned data into json
        data_dict = response.json()

        # get the data list
        data_list = data_dict['data']['home_search']['results']

        page +=1
        print("scraping www.realtor.com, Now is page",page,"..Please wait..")

        # offset increase 200 each iteration
        offset += 200

        # if the data_list has length 0，it reached the last page. return and generate the excel workbook.
        if len(data_list) == 0:
            print('Last page reached, completed!')
            workbook.save('data.xls')               # Save the excel with name data.xls
            break

        else:         
            for data in data_list:
                #print(data)
                
                # if list_price is none, then use most recent sold_price
                if data['list_price'] == None:
                    sheet.write(row, 0, data['description']['sold_price'])
                else:
                    sheet.write(row, 0, data['list_price'])

                # if list_date is none, then use most recent sold_date
                if data['list_date'] == None:
                    sheet.write(row, 4, data['description']['sold_date'])
                else:
                    sheet.write(row, 4, data['list_date'][:10])

                # Write into excel.
                sheet.write(row, 1, data['description']['beds'])
                sheet.write(row, 2, data['description']['baths'])
                sheet.write(row, 3, data['description']['sqft'])
                sheet.write(row, 5, data['description']['type'])
                sheet.write(row, 6, data['branding'][0]['name'])
                sheet.write(row, 7, data['location']['address']['line'])
                sheet.write(row, 8, data['location']['address']['city'])
                sheet.write(row, 9, data['location']['address']['state'])
                sheet.write(row, 10, data['location']['address']['postal_code'])
                # row by row
                row += 1

#Main
if __name__ == '__main__':
    # cookies data.
    # Allow servers to communicate using a small piece of data. Alao used to identify if the reqeust is coming from a real user or a bot
    cookies = {
        'split': 'n',
        'split_tcv': '173',
        '__vst': '0e04d5de-20ea-4926-bcf6-49428c17dc97',
        '__ssn': 'd9c3a74f-9561-4ed7-9bfd-b19c6127472b',
        '__ssnstarttime': '1669260422',
        'permutive-id': 'b1870170-7f3d-4884-809f-d381c69884fa',
        '_pbjs_userid_consent_data': '3524755945110770',
        'pxcts': 'df543e7c-6ba7-11ed-8cf0-6763476b7867',
        '_pxvid': 'df54302f-6ba7-11ed-8cf0-6763476b7867',
        '__split': '38',
        'AMCVS_8853394255142B6A0A4C98A4%40AdobeOrg': '1',
        's_ecid': 'MCMID%7C47862618717714335251426558045369882772',
        '__gads': 'ID=69b306cf834dfb1f:T=1669260428:S=ALNI_MYIOHZ5kzuLfY7joG0NroPezTcyOA',
        '__gpi': 'UID=00000b8240a9bbd7:T=1669260428:RT=1669260428:S=ALNI_MbinAnCmfe_UvXGsLl0adoZQgBMhA',
        'AMCV_8853394255142B6A0A4C98A4%40AdobeOrg': '-1124106680%7CMCIDTS%7C19321%7CMCMID%7C47862618717714335251426558045369882772%7CMCAAMLH-1669865228%7C11%7CMCAAMB-1669865228%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1669267629s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.2.0',
        'ajs_anonymous_id': '%22293fe8b2-1854-4146-8740-cbf4e547bb2c%22',
        '_gcl_au': '1.1.1138838367.1669260430',
        '_ncg_id_': 'c7cfa266-2ff2-4d39-84c4-00988ea40bfb',
        '_tac': 'false~self|not-available',
        '_ta': 'us~1~89df68bc35032d48c293adce8574aa42',
        '_ncg_domain_id_': 'c7cfa266-2ff2-4d39-84c4-00988ea40bfb.1.1669260430.1732332430',
        '_lr_geo_location': 'CN',
        '_gid': 'GA1.2.573563840.1669260432',
        'AMCVS_AMCV_8853394255142B6A0A4C98A4%40AdobeOrg': '1',
        'AMCV_AMCV_8853394255142B6A0A4C98A4%40AdobeOrg': '-1124106680%7CMCMID%7C47862618717714335251426558045369882772%7CMCIDTS%7C19321%7CMCOPTOUT-1669267633s%7CNONE%7CvVersion%7C5.2.0',
        '__qca': 'P0-141185221-1669260430298',
        '_ncg_g_id_': '2df7f917-9a34-4b26-af23-922f237dbcfd.3.1669260431.1732332430',
        '_lr_sampling_rate': '100',
        '_lr_env_src_ats': 'false',
        'ab.storage.deviceId.7cc9d032-9d6d-44cf-a8f5-d276489af322': '%7B%22g%22%3A%22c3e4eb4b-e3da-b5bb-dd30-63fa167bac51%22%2C%22c%22%3A1669260549861%2C%22l%22%3A1669260549927%7D',
        'ab.storage.userId.7cc9d032-9d6d-44cf-a8f5-d276489af322': '%7B%22g%22%3A%22visitor_0e04d5de-20ea-4926-bcf6-49428c17dc97%22%2C%22c%22%3A1669260549903%2C%22l%22%3A1669260549931%7D',
        '_iidt': '68mdwbNA8nSsFu3POB90jQ18QXkW5furACObKfAqVWuInZZBMUuzZZq4YA857CfU/3oWgtPumVdQ6w==',
        '_vid_t': 'MYgFJRp4MrgFJPC3VbRWANdR+Jby89dd5t+JLciSahLuuNmYVwBdI+jr7sTJCFokxBkSpn8uRQBWMw==',
        '__fp': 'lcV8MSByFgDYKKVhafcW',
        'user_activity': 'return',
        'last_ran': '-1',
        '__opTest': '3',
        'QSI_HistorySession': 'https%3A%2F%2Fwww.realtor.com%2F~1669260551741%7Chttps%3A%2F%2Fwww.realtor.com%2Frealestateandhomes-search%2FSan-Diego_CA~1669261123481',
        'srchID': '0e1adcda59e84fdc95c5870cd3398d0f',
        '_ncg_sp_id.cc72': 'c7cfa266-2ff2-4d39-84c4-00988ea40bfb.1669260430.1.1669261388.1669260430.6ab723da-eb2b-4a29-8f15-ed7552e1cecb',
        '_ga': 'GA1.2.2088485266.1669260429',
        'adcloud': '{%22_les_v%22:%22y%2Crealtor.com%2C1669263239%22}',
        'cto_bundle': 'aXWKz19uYmZ5djg2OXclMkJOS0YxajNkdjhrUmRGU0htZjM0UkkxZWhIQW50anptRjYlMkJ5UEF1QlNrRUpnbWl2SDZPamFhYnB6aThPUWllJTJCZVl2YU85JTJGY2xwS2FvZUcwUWE1bFBOSFJqaFBIaDlkVnBtU2VUU3E3VElHVld3VSUyQlNNd0JSY2xIQVZZVmtYZ21YNlhFRW5vTjNVdXZBJTNEJTNE',
        '_uetsid': 'e2ac7e106ba711ed90e045701c032c98',
        '_uetvid': 'e2ad20a06ba711edb540f7c53fd24e6e',
        'ab.storage.sessionId.7cc9d032-9d6d-44cf-a8f5-d276489af322': '%7B%22g%22%3A%225abf9523-6db6-f642-c1cd-cfedfd1920a2%22%2C%22e%22%3A1669264140183%2C%22c%22%3A1669260549923%2C%22l%22%3A1669262340183%7D',
        '_tas': 'mezajnsn0tt',
        'srchID': 'f4a4b5f5e3974a06865bc9466a7137f1',
        'criteria': 'pg%3D57%26sprefix%3D%252Frealestateandhomes-search%26area_type%3Dcity%26search_type%3Dcity%26city%3DSan%2520Diego%26state_code%3DCA%26state_id%3DCA%26lat%3D32.814977%26long%3D-117.1355615%26county_fips%3D06073%26county_fips_multi%3D06073%26loc%3DSan%2520Diego%252C%2520CA%26locSlug%3DSan-Diego_CA%26county_needed_for_uniq%3Dfalse',
        '_px3': '406fbbb6dc2a6b13add431d03b9f64b7b2acd08594144739dddd23f18a9e3787:opUkgz9y9lGe0jK5W/f3QuXhnOAKgGLoZlHGYbqT5qnecJ1I9sisEZhlbzXZp9vj5xyMrR2b8U2DHXSBPTuk6A==:1000:vTttEc3u/XiSmo7v1uUxj3DRXZPDFrNilWXSZSaVcsNxO7Db74o9XuQU0DopkPx1MnJoLh7JX3V5Dv79XFcX6T3EZOoHYz4Fa9knAhTPdhdQj8zJK5mqYP4A6h+iBnBABCfJ3tfMRYqlEni5Vq94gHCI8dv5+e83pW0KC01wzqbf1fXNEsJJ8FSf0nKy4rgMr1p4zOFbIWSN+GdkqbDxVQ==',
        '_ga_MS5EHT6J6V': 'GS1.1.1669260431.1.1.1669264167.37.0.0',
        'last_ran_threshold': '1669264188633',
    }

    #headers data.
    #Used to passes additional context and metadata about the request or response.
    headers = {
        'authority': 'www.realtor.com',
        'accept': 'application/json',
        'accept-language': 'en-US,en;q=0.9',
        # Already added when you pass json=
        # 'content-type': 'application/json',
        # Requests sorts cookies= alphabetically
        # 'cookie': 'split=n; split_tcv=173; __vst=0e04d5de-20ea-4926-bcf6-49428c17dc97; __ssn=d9c3a74f-9561-4ed7-9bfd-b19c6127472b; __ssnstarttime=1669260422; permutive-id=b1870170-7f3d-4884-809f-d381c69884fa; _pbjs_userid_consent_data=3524755945110770; pxcts=df543e7c-6ba7-11ed-8cf0-6763476b7867; _pxvid=df54302f-6ba7-11ed-8cf0-6763476b7867; __split=38; AMCVS_8853394255142B6A0A4C98A4%40AdobeOrg=1; s_ecid=MCMID%7C47862618717714335251426558045369882772; __gads=ID=69b306cf834dfb1f:T=1669260428:S=ALNI_MYIOHZ5kzuLfY7joG0NroPezTcyOA; __gpi=UID=00000b8240a9bbd7:T=1669260428:RT=1669260428:S=ALNI_MbinAnCmfe_UvXGsLl0adoZQgBMhA; AMCV_8853394255142B6A0A4C98A4%40AdobeOrg=-1124106680%7CMCIDTS%7C19321%7CMCMID%7C47862618717714335251426558045369882772%7CMCAAMLH-1669865228%7C11%7CMCAAMB-1669865228%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1669267629s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.2.0; ajs_anonymous_id=%22293fe8b2-1854-4146-8740-cbf4e547bb2c%22; _gcl_au=1.1.1138838367.1669260430; _ncg_id_=c7cfa266-2ff2-4d39-84c4-00988ea40bfb; _tac=false~self|not-available; _ta=us~1~89df68bc35032d48c293adce8574aa42; _ncg_domain_id_=c7cfa266-2ff2-4d39-84c4-00988ea40bfb.1.1669260430.1732332430; _lr_geo_location=CN; _gid=GA1.2.573563840.1669260432; AMCVS_AMCV_8853394255142B6A0A4C98A4%40AdobeOrg=1; AMCV_AMCV_8853394255142B6A0A4C98A4%40AdobeOrg=-1124106680%7CMCMID%7C47862618717714335251426558045369882772%7CMCIDTS%7C19321%7CMCOPTOUT-1669267633s%7CNONE%7CvVersion%7C5.2.0; __qca=P0-141185221-1669260430298; _ncg_g_id_=2df7f917-9a34-4b26-af23-922f237dbcfd.3.1669260431.1732332430; _lr_sampling_rate=100; _lr_env_src_ats=false; ab.storage.deviceId.7cc9d032-9d6d-44cf-a8f5-d276489af322=%7B%22g%22%3A%22c3e4eb4b-e3da-b5bb-dd30-63fa167bac51%22%2C%22c%22%3A1669260549861%2C%22l%22%3A1669260549927%7D; ab.storage.userId.7cc9d032-9d6d-44cf-a8f5-d276489af322=%7B%22g%22%3A%22visitor_0e04d5de-20ea-4926-bcf6-49428c17dc97%22%2C%22c%22%3A1669260549903%2C%22l%22%3A1669260549931%7D; _iidt=68mdwbNA8nSsFu3POB90jQ18QXkW5furACObKfAqVWuInZZBMUuzZZq4YA857CfU/3oWgtPumVdQ6w==; _vid_t=MYgFJRp4MrgFJPC3VbRWANdR+Jby89dd5t+JLciSahLuuNmYVwBdI+jr7sTJCFokxBkSpn8uRQBWMw==; __fp=lcV8MSByFgDYKKVhafcW; user_activity=return; last_ran=-1; __opTest=3; QSI_HistorySession=https%3A%2F%2Fwww.realtor.com%2F~1669260551741%7Chttps%3A%2F%2Fwww.realtor.com%2Frealestateandhomes-search%2FSan-Diego_CA~1669261123481; srchID=0e1adcda59e84fdc95c5870cd3398d0f; _ncg_sp_id.cc72=c7cfa266-2ff2-4d39-84c4-00988ea40bfb.1669260430.1.1669261388.1669260430.6ab723da-eb2b-4a29-8f15-ed7552e1cecb; _ga=GA1.2.2088485266.1669260429; adcloud={%22_les_v%22:%22y%2Crealtor.com%2C1669263239%22}; cto_bundle=aXWKz19uYmZ5djg2OXclMkJOS0YxajNkdjhrUmRGU0htZjM0UkkxZWhIQW50anptRjYlMkJ5UEF1QlNrRUpnbWl2SDZPamFhYnB6aThPUWllJTJCZVl2YU85JTJGY2xwS2FvZUcwUWE1bFBOSFJqaFBIaDlkVnBtU2VUU3E3VElHVld3VSUyQlNNd0JSY2xIQVZZVmtYZ21YNlhFRW5vTjNVdXZBJTNEJTNE; _uetsid=e2ac7e106ba711ed90e045701c032c98; _uetvid=e2ad20a06ba711edb540f7c53fd24e6e; ab.storage.sessionId.7cc9d032-9d6d-44cf-a8f5-d276489af322=%7B%22g%22%3A%225abf9523-6db6-f642-c1cd-cfedfd1920a2%22%2C%22e%22%3A1669264140183%2C%22c%22%3A1669260549923%2C%22l%22%3A1669262340183%7D; _tas=mezajnsn0tt; srchID=f4a4b5f5e3974a06865bc9466a7137f1; criteria=pg%3D57%26sprefix%3D%252Frealestateandhomes-search%26area_type%3Dcity%26search_type%3Dcity%26city%3DSan%2520Diego%26state_code%3DCA%26state_id%3DCA%26lat%3D32.814977%26long%3D-117.1355615%26county_fips%3D06073%26county_fips_multi%3D06073%26loc%3DSan%2520Diego%252C%2520CA%26locSlug%3DSan-Diego_CA%26county_needed_for_uniq%3Dfalse; _px3=406fbbb6dc2a6b13add431d03b9f64b7b2acd08594144739dddd23f18a9e3787:opUkgz9y9lGe0jK5W/f3QuXhnOAKgGLoZlHGYbqT5qnecJ1I9sisEZhlbzXZp9vj5xyMrR2b8U2DHXSBPTuk6A==:1000:vTttEc3u/XiSmo7v1uUxj3DRXZPDFrNilWXSZSaVcsNxO7Db74o9XuQU0DopkPx1MnJoLh7JX3V5Dv79XFcX6T3EZOoHYz4Fa9knAhTPdhdQj8zJK5mqYP4A6h+iBnBABCfJ3tfMRYqlEni5Vq94gHCI8dv5+e83pW0KC01wzqbf1fXNEsJJ8FSf0nKy4rgMr1p4zOFbIWSN+GdkqbDxVQ==; _ga_MS5EHT6J6V=GS1.1.1669260431.1.1.1669264167.37.0.0; last_ran_threshold=1669264188633',
        'newrelic': 'eyJ2IjpbMCwxXSwiZCI6eyJ0eSI6IkJyb3dzZXIiLCJhYyI6IjM3ODU4NCIsImFwIjoiMTI5NzQxMzUyIiwiaWQiOiJhNWVmZjI4NGZmMzA1M2UxIiwidHIiOiI5ZmU0NzdlYTgzZWQ4ODk1MzY4MjA3NzdlYzU5ZWZkYiIsInRpIjoxNjY5MjY0MTg4NjQ0LCJ0ayI6IjEwMjI2ODEifX0=',
        'origin': 'https://www.realtor.com',
        'referer': 'https://www.realtor.com/realestateandhomes-search/San-Diego_CA/pg-57',
        'sec-ch-ua': '"Google Chrome";v="107", "Chromium";v="107", "Not=A?Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'traceparent': '00-9fe477ea83ed889536820777ec59efdb-a5eff284ff3053e1-01',
        'tracestate': '1022681@nr=0-1-378584-129741352-a5eff284ff3053e1----1669264188644',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36',
    }

    #Client_id and Schema
    params = {
        'client_id': 'rdc-x',
        'schema': 'vesta',
    }

    # Names of the input GUI
    fieldNames = ['Ex : Urbana, IL' ]
    msg = 'Input Location Name(City, State)'
    title = 'CS410 Final Project'

    # get the values from the input gui that entered by users
    fieldValues = gui.multenterbox(msg, title, fieldNames)
    
    #initilize the offset to 0
    offset = 0

    #Limit to return 200 rows per page
    limit = 200

    #get the city name
    location = fieldValues[0]

    # Initilzie the row in excel to 1
    row = 1

    # genereate excel
    workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
    
    # generte work sheet
    sheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    
    # write the column names
    sheet.write(0, 0, 'price')
    sheet.write(0, 1, 'beds_num')
    sheet.write(0, 2, 'baths_num')
    sheet.write(0, 3, 'sqft')
    sheet.write(0, 4, 'list_date')
    sheet.write(0, 5, 'type')
    sheet.write(0, 6, 'company_name')
    sheet.write(0, 7, 'street')
    sheet.write(0, 8, 'city')
    sheet.write(0, 9, 'state')
    sheet.write(0, 10, 'postal_code')

    #Execute method
    run(offset=0, limit=200, location=location, row=1)








