# -*- coding: utf-8 -*-
import urllib
import urllib2
from bs4 import BeautifulSoup
import os

# link to argentinian energy department sworn statements
BASE_URL = 'https://www.se.gob.ar/datosupstream/consulta_avanzada/ddjj.php'
REQ_URL = "https://www.se.gob.ar/datosupstream/consulta_avanzada/ddjj.xls.php"


def scrape_id_empresa():
    """Scrape key values of "idempresa" field."""

    # get html
    html = urllib2.urlopen(BASE_URL)

    # make beautiful soup with html
    bs = BeautifulSoup(html)

    # find tags containing "idempresa"
    options_list = bs.find("select", {"name": "idempresa"}).find_all("option")

    # extract key values from each tag
    values_list = [option["value"] for option in options_list if
                   option["value"] != ""]

    return values_list


def download_ddjj(company, year, month, base_path):
    """Download one ddjj of a certain company-year-month."""

    # form data to make the request
    values = {"idempresa": company,
              "idmes": month,
              "idanio": year,
              "submit": "Bajar+Excel"
              }

    # encode values to make the request
    data = urllib.urlencode(values)

    # build request
    req = urllib2.Request(REQ_URL, data)

    # send request
    response = urllib2.urlopen(req)

    # create file name
    file_name = str(company) + "_" + str(year) + "_" + str(month) + ".xls"

    # create base path if not exists
    if not os.path.isdir(base_path):
        os.makedirs(base_path)

    # file complete path
    file_path = os.path.join(base_path, file_name)

    # save it
    with open(file_path, "wb") as local_file:
        local_file.write(response.read())


def download_all_ddjj(base_path):
    """Download all ddjjs."""

    # take key values to build queries
    companies = scrape_id_empresa()
    years = range(2006, 2015)
    months = range(1, 13)

    # calculate progress indicator
    num_files = len(companies) * len(years) * len(months)

    # download all files
    counter = 1
    for company in companies:
        for year in years:
            for month in months:

                file_name = str(company) + "_" + str(year) + "_" + \
                    str(month) + ".xls"

                print "Downloading " + file_name,

                download_ddjj(company, year, month, base_path)

                print str(counter) + " of " + str(num_files) + \
                    " files downloaded."

                # increment counter
                counter += 1


def main(output_path=None):

    # take path passed or default value
    base_path = output_path or "ddjj"

    # download all ddjj from argentinian energy department
    download_all_ddjj(base_path)


if __name__ == '__main__':

    if len(sys.argv) == 2:
        main(sys.argv[1])

    else:
        main()
