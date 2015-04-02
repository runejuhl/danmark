#!/usr/bin/env python3

import requests
import xlrd
import os, sys

import json, yaml

def run():
    f = '../files/postdk-post_codes.xls'
    x = None
    try:
        os.stat(f)
    except FileNotFoundError as e:
        r = requests.get('https://www.postdanmark.dk/da/Documents/Lister/postnummerfil-excel.xls')
        open(f, 'wb').write(r.content)

    # open xls
    x = xlrd.open_workbook(f)
    # open first sheet
    s = x.sheets()[0]

    # first row is date, second row is headings
    assert(s.row(1)[0].value == 'Postnr.')
    assert(s.row(1)[1].value == 'Bynavn')

    data = {}
    for i in range(2, s.nrows):
        data[int(s.row(i)[0].value)] = s.row(i)[1].value.strip()

    json.dump(data, open('../output/postdk-post_codes.json', 'w'))
    open('../output/postdk-post_codes.yaml', 'w').write(yaml.dump(data, default_flow_style=False))

if __name__ == '__main__':
    run()
