from pyExcelerator.ImportXLS import parse_xls

TERM_WIDTH = 80

def xls_to_excelerator_dict(filename, encoding=None):
    return parse_xls(filename, encoding)

def print_excelerator_dict(excelerator_dict):
    for sheet_name, data in excelerator_dict:
        print u'Sheet: %s' % sheet_name
        for k, v in sorted(data.items()):
            print k, v

def sequence_with_index(seq):
    return zip(range(len(seq)), seq)

def print_matrix(matrix):
    for sheet, rows in matrix.iteritems():
        print 'Sheet: %s' % sheet

        for (row_num, columns) in sequence_with_index(rows):
            print 'row: %d' % row_num
            for (col_num, v) in sequence_with_index(columns):
                print col_num, v
            print

def excelerator_dict_to_rows_and_columns(excelerator_dict):
    sheets = {}

    for sheet, data in excelerator_dict:
        rows = {}
        for ((row_num, col_num), v) in data.iteritems():
            rows.setdefault(row_num, {})[col_num] = v

        sheets[sheet] = rows

    return sheets

def rows_and_columns_to_matrix(rows_and_columns):
    new_sheets = {}
    for (sheet, d) in rows_and_columns.iteritems():
        new_sheet = []
        row_numbers = xrange(max(d.keys()) + 1)
        for row_number in row_numbers:
            row = d.get(row_number)
            col_numbers = xrange(max(row.keys()) + 1)
            new_row = [row.get(i) for i in col_numbers]

            new_sheet.append(new_row)

        new_sheets[sheet] = new_sheet

    return new_sheets

def sequence_to_dict(seq, index_names):
    return dict(zip(index_names, seq))

def matrix_to_structure(matrix):
    return dict(
        (sheet, [sequence_to_dict(cols, rows[0]) for cols in rows[1:]])
        for (sheet, rows) in matrix.iteritems()
    )

def print_dict(d):
    for item in d.iteritems():
        print '%s: %s' % item

def print_structure(dicts):
    for sheet, data in dicts.iteritems():
        print 'sheet: %s' % sheet
        for d in data:
            print_dict(d)
            print

def xls_to_structure(path):
    return matrix_to_structure(
        rows_and_columns_to_matrix(
            excelerator_dict_to_rows_and_columns(
                xls_to_excelerator_dict(path)
            )
        )
    )


