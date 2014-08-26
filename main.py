# coding: utf-8

import argparse
import logging
import xlrd


def get_db_rows(db_xlsx_filename):
  """Return a list of rows from the XLSX file.

  Each row is a dict.  Keys are column names (e.g., NAME) and values are the
  cells in the row (e.g., Lavender Lemon).
  """
  workbook = xlrd.open_workbook(db_xlsx_filename)
  db_worksheet = workbook.sheet_by_name('Sheet1')
  column_names = [value for value in db_worksheet.row_values(0) if value]
  column_count = len(column_names)
  rows = []
  for row_index in range(1, db_worksheet.nrows):
    values = [value for value in db_worksheet.row_values(row_index)]
    values = values[:column_count]
    # Skip rows that have no values
    if not len(filter(None, values)):
      continue
    rows.append(dict(zip(column_names, values)))
  return rows


def export_to_taplister(beers):
  """Export the ``beers`` to taplister."""
  # TODO: implement
  pass


def export_to_text(beers, text_filename):
  """Export the ``beers`` to a text file."""
  # TODO: implement
  pass


def main(args):
  beers = get_db_rows(args.xlsx_file)
  for beer in beers:
    name = beer['NAME']
    price = beer['PRICE']
    print name, price


def process_args():
  parser = argparse.ArgumentParser()
  parser.add_argument('--xlsx-file', default='taplistDB.xlsx',
                      help='The XLSX file containing the tap list')
  parser.add_argument('--debug',
                      action='store_true', default=False,
                      help='Display debug messages')
  args = parser.parse_args()
  if args.debug:
    log.setLevel(logging.DEBUG)
    logging.getLogger('root').setLevel(logging.DEBUG)
  return args


if __name__ == '__main__':
 main(process_args())
