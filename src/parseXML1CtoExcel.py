#!/usr/bin/python
import argparse
import pickle
import re
import xml.sax
import hashlib

# import redis
import xlsxwriter


class MyHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.tags = []
        self.cnt = 0
        self.data = {}
        self.itemName = ''
        self.item = {}
        self.currentTag = ''

    def startElement(self, tag, attributes):
        self.cnt += 1
        if self.cnt % 4096 == 0: print(self.cnt)

        self.tags.append(tag)

        if len(self.tags) == 3:
            self.itemName = tag
            self.item = {}
            self.currentTag = ''
            pass
        elif len(self.tags) == 4:
            self.currentTag = tag

    def endElement(self, tag):
        self.tags.pop()
        if len(self.tags) == 2:
            if self.itemName not in self.data.keys(): self.data[self.itemName] = []
            self.data[self.itemName].append(dict(self.item))

    @staticmethod
    def normalizeString(x):
        x = x.strip()
        if re.match('^[0-]$', x): x = ''
        return x

    def characters(self, content):
        if self.currentTag != '':
            if self.currentTag not in self.item.keys(): self.item[self.currentTag] = ''
            self.item[self.currentTag] += self.normalizeString(content)


if __name__ == "__main__":
    argparser = argparse.ArgumentParser(description='Parse 1C XML export file')
    argparser.add_argument('input_file', help='input XML-file')
    argparser.add_argument('output_dir', help='output folder')
    argparser.add_argument('--cache', dest='cache_file', help='cache file path')
    args = argparser.parse_args()

    data = None

    if args.cache_file is not None:
        try:
            print("load cache from %s" % args.cache_file)
            with open(args.cache_file, 'rb') as f:
                data = pickle.load(f)
        except FileNotFoundError:
            print("can't find cache %s" % args.cache_file)

    if data is None:
        parser = xml.sax.make_parser()
        parser.setFeature(xml.sax.handler.feature_namespaces, 0)
        print("parse file %s" % args.input_file)
        handler = MyHandler()
        parser.setContentHandler(handler)
        parser.parse(args.input_file)
        print('items count: %d. End parsing' % handler.cnt)
        data = handler.data
        if args.cache_file is not None:
            print("save cache to %s" % args.cache_file)
            with open(args.cache_file, 'wb') as f:
                pickle.dump(data, f)

    for sheetName in data.keys():
        lenSheetName = len(data[sheetName])
        columns = {}
        col = 0
        sheetData = data[sheetName]
        rowCount = len(sheetData)

        for rowDict in sheetData:
            for colName in rowDict.keys():
                if colName not in columns.keys():
                    columns[colName] = col
                    col += 1
        print("sheet '%s': %s, count: %d" % (sheetName, columns, len(sheetData)))

        book = xlsxwriter.Workbook("%s/%s.xlsx" % (args.output_dir, sheetName))
        sheet = book.add_worksheet()
        row = 0
        for colName in columns.keys():
            sheet.write(row, columns[colName], colName)
        for rowDict in sheetData:
            if row % 128 == 0: print('save %s: %d / %d' % (sheetName, row, rowCount))
            row += 1
            for colName in rowDict.keys():
                sheet.write(row, columns[colName], rowDict[colName])
        print('save %s: %d / %d' % (sheetName, row, rowCount))
        book.close()
