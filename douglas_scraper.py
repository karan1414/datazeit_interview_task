import json
import re
from pprint import pprint

from requests_html import HTMLSession
from xlsxwriter import Workbook

# URLS Required
home_url = 'https://www.douglas.de/de'
product_base_url = 'https://www.douglas.de'
products_page_api = '/de/c/gesicht/gesichtsmasken/feuchtigkeitsmasken/120308'
product_info_url = 'https://www.douglas.de/api/v2/products/'


# XPATHS Required

# 1. Xpath for product pages
product_pages_xpath = '//a[@class="link link--no-decoration pagination-title__option-link active"]/@href'
# 2. Xpath For products <a> tag
product_a_tag_xpath = '//a[@class="link link--no-decoration product-tile__main-link"]/@href'

# Regex Required
product_number_regex = r'\s*\/[a-z]+\/[a-z]+\/(.*)'

# Creating a session
session = HTMLSession()

# paramaters required for requesting in product page
product_page_params = {"fields": "FULL"}

# To display headers and subsequent product data in this sequence
EXCEL_HEADERS = ['ean', 'product_name', 'product_description', 'product_details', 'product_features', 'variant_name', 'product_image_link', 'final_price', 'original_price', 'discount_percentage', 'rating', 'number_of_reviews', 'is_available' ]

def write_product_details_to_excel(product_details):
    ''' 
    This function takes in product details array containing product details extracted from the product_info_json_arr and write them into excel file. If the product is out of stock then the row is highlighted by red color.
    '''

    wb = Workbook('douglas_product_details.xlsx')
    ws = wb.add_worksheet('gesichtsmasken_products')
    

    out_of_stock = wb.add_format({'bg_color': '#FFC7CE'})
    header_format = wb.add_format({'bold': True, 'border': 1})
    column_format = wb.add_format({'text_wrap': True})
    column_format.set_align('justify')
    column_format.set_align('vjustify')

    first_row = 0
    # workbook headers
    for header in EXCEL_HEADERS:
        column = EXCEL_HEADERS.index(header)
        ws.write(first_row, column, header, header_format)

    row = 1
    for product in product_details:
        for key,value in product.items():
            column = EXCEL_HEADERS.index(key)
            if key == "is_available" and value == False:
                ws.set_row(row, cell_format=out_of_stock)
            if key == "product_features":
                value = ",\n".join(value)
            if key == "product_details":
                final_value = ""
                for detail in product["product_details"]:
                    for k,v in detail.items():
                        final_value += "{} - {} \n".format(k, v)
                value = final_value
            ws.write(row, column, str(value))
        row += 1
    ws.set_column("A:M", 30, cell_format=column_format)
    wb.close()
    return True

def get_product_details(product_info_json_arr):
    ''' This function recieves the product info jsons in a list and iterate over them and extract the details required and adds to product detail dictionary and adds this dictionary to product_details_arr which is then returned '''
    product_details_arr = []
    for product_info_json in product_info_json_arr:
        product_detail = {}   

        # isAvailable(is in stock)
        if 'stock' in product_info_json and product_info_json['stock']:
            product_detail['is_available'] = True if 'stockLevel' in product_info_json['stock'] and product_info_json['stock']['stockLevel'] else False

        # ean unique number
        if "ean" in product_info_json and product_info_json['ean']:
            product_detail['ean'] = product_info_json["ean"]
        
        # product name
        if 'brandLine' in product_info_json and product_info_json['brandLine'] and 'name' in product_info_json['brandLine'] and product_info_json['brandLine']['name'] and 'baseProductName' in product_info_json and product_info_json['baseProductName']:
                product_detail['product_name'] = product_info_json['brandLine']['name'] + product_info_json['baseProductName']  

        # product description
        if "description" in product_info_json and product_info_json["description"]:
            product_detail['product_description'] = product_info_json["description"]

        # ratings
        if "ratingStars" in product_info_json and product_info_json["ratingStars"]:
            product_detail['rating'] = product_info_json["ratingStars"]
    
        # number of reviews
        if "numberOfReviews" in product_info_json and product_info_json["numberOfReviews"]:
            product_detail['number_of_reviews'] = product_info_json["numberOfReviews"]

        # variant name
        if "name" in product_info_json and product_info_json["name"]:
            product_detail['variant_name'] = product_info_json["name"]

        # price
        if 'price' in product_info_json and product_info_json["price"] and "formattedValue" in product_info_json["price"] and product_info_json["price"]["formattedValue"]:
            product_detail['final_price'] = product_info_json["price"]["formattedValue"].replace("\xa0", "")

            if "formattedOriginalValue" in product_info_json["price"] and product_info_json["price"]["formattedOriginalValue"]:
                product_detail['original_price'] = product_info_json["price"]["formattedOriginalValue"]

            if "discountPercentage" in product_info_json["price"] and product_info_json["price"]["discountPercentage"]:
                product_detail["discount_percentage"] = product_info_json["price"]["discountPercentage"]

        # image url
        if 'productApplicationImage' in product_info_json and product_info_json["productApplicationImage"] and product_info_json["productApplicationImage"] and "url" in product_info_json["productApplicationImage"] and product_info_json["productApplicationImage"]["url"]:
            product_detail['product_image_link'] = product_info_json["productApplicationImage"]["url"]

        # product featuers and details
        if 'classifications' in product_info_json and product_info_json["classifications"]:
            product_labels = []
            product_features = []
            product_details = []
            for feature in product_info_json["classifications"]:
                if "features" not in feature or not feature["features"]:
                    continue

                for fv in feature["features"]:
                    if "productLabel" in fv and fv["productLabel"]:
                        product_labels = [f["value"] for f in fv["productLabel"]]
                    values = [] 
                    k = fv["name"]
                    values = [ft_value["value"] for ft_value in fv["featureValues"]]
                    v = ",".join(values)
                    product_details.append({k: v})
            product_detail["product_features"] = product_labels
            product_detail["product_details"] = product_details
        product_details_arr.append(product_detail)
    return product_details_arr

def get_products(products_page_api):
    ''' This function takes in the product page api and returns all the product a tags 
    product_page_api = Api of the product page.
    '''
    prods = []
    products_page_url = product_base_url + products_page_api

    try:
        products_page_response = session.get(products_page_url)
    except Exception as e:
        print('Error-while-getting-product-page-{}'.format(e))
        return prods
    products_page_response.html.render(sleep=1)
    
    prods = products_page_response.html.xpath(product_a_tag_xpath)    

    return prods

def get_product_info_arr(products):
    ''' This function takes in list of all the products href and makes get request to the same and add the required values from resultant json to dictionary and append the dictonary to result list and  returns product information in list of dictionary format.
    '''
    result_list_product_info = []
    for product in products:
        product_groups = re.search(product_number_regex, product, re.I)
        if not product_groups or not product_groups.groups(1):
            continue
        product_number = product_groups.group(1)
        product_details_url = product_info_url + product_number
        product_page_resp = session.get(product_details_url, params=product_page_params)
        product_info_json = json.loads(product_page_resp.text)
        if not product_info_json:
            continue
        result_list_product_info.append(product_info_json)
    return result_list_product_info

def scrapeProductPage():
    ''' This function establishes a session for the given page and makes requests to reach the product details page and stores the json response of product detail page to process it further '''
    products = []
    try:
        homepage_response = session.get(home_url)
    except Exception as e:
        print('Error-while-getting-homepage-{}'.format(e))

    products_page_url = product_base_url + products_page_api

    try:
        products_page_response = session.get(products_page_url)
    except Exception as e:
        print('Error-while-getting-product-page-{}'.format(e))
    
    products_page_response.html.render(sleep=1)
    
    products = products_page_response.html.xpath(product_a_tag_xpath)
    product_pages = products_page_response.html.xpath(product_pages_xpath)
    for prod_page in product_pages:
        products_found = get_products(prod_page)
        if not products_found:
            print("unable-to-find-products-for-{}".format(prod_page))
            continue
        products.extend(products_found)

    product_info_json_arr = get_product_info_arr(products)
    if not product_info_json_arr:
        print("unable-to-create-product_info_json_arr")
        return

    product_details = get_product_details(product_info_json_arr)
    if not product_details:
        print("unable-to-create-product_details")
        return
 
    write_product_details_to_excel(product_details)


if __name__ == '__main__':
    scrapeProductPage()
