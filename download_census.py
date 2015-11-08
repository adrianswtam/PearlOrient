#!/usr/bin/env python

# Download Census data from census2011.gov.hk

import urllib2
import re
import os

def getDistrictCodes():
    MAINPAGE='http://www.census2011.gov.hk/en/constituency-area.html'
    AGENT='Mozilla/5.0 (Windows; U; Windows NT 5.1; it; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11'
    for line in urllib2.urlopen(urllib2.Request(MAINPAGE, headers={'User-Agent':AGENT})):
        m = re.search(r'"\(([a-z]\d{1,2})\) (.*)": "\1"', line, re.I)
        if m:
            yield m.group(1), m.group(2)

def getExcelUrls():
    return sorted(set('http://idds.census2011.gov.hk/Fact_sheets/CA/%s.xlsx' % code for code, _ in getDistrictCodes()))

def downloadExcels():
    try:
        os.makedirs("_census_ca")
    except:
        pass # ignore for dir already created
    for url in getExcelUrls():
        filename = os.path.join("_census_ca",url.rsplit('/')[-1])
        if os.path.isfile(filename):
            continue # skip if file exists
        resp = urllib2.urlopen(url, timeout=5)
        content = resp.read()
        print "Saving %s as %s (%d bytes)" % (url, filename, len(content))
        with open(filename, "wb") as fp:
            fp.write(content)

if __name__ == '__main__':
    downloadExcels()

# vim:set ts=4 sw=4 sts=4 et:
