import pandas as pd
import argparse
import sys

if __name__ == '__main__':

    # parse command-line arguments using argparse
    parser = argparse.ArgumentParser(
        prog='diffxl',
        description='Describe the diff of 2 excel files(Only values in each cell. Texts in shapes are ignored).'
    )
    parser.add_argument(
        'path1',
        metavar='path1',
        nargs=1,
        help='A path to your excel file'
    )
    parser.add_argument(
        'path2',
        metavar='path2',
        nargs=1,
        help='A path to your excel file'
    )

    args = parser.parse_args()
    path1 = args.path1[0]
    path2 = args.path2[0]

    # Read the files
    # Exit if the file doesn't exist
    f1 = None
    try:
        f1 = pd.read_excel(path1, sheet_name=None, header=None)
    except FileNotFoundError:
        print(f'The file \'{path1}\' is not found')
        sys.exit(1)

    f2 = None
    try:
        f2 = pd.read_excel(path2, sheet_name=None, header=None)
    except FileNotFoundError:
        print(f'The file \'{path2}\' is not found')
        sys.exit(1)

    # Compare the sheet names of 2 files
    # and make the list of sheet names which exist in both files
    print(f'Sheet-level comparision between \'{path1}\' and \'{path2}\'...')
    sheets = set()
    for f1Sheet in list(f1):
        if not f1Sheet in f2:
            print(f'The sheet \'{f1Sheet}\' doesn\'t exist in {path2}')
        else:
            sheets.add(f1Sheet)

    for f2Sheet in list(f2):
        if not f2Sheet in f1:
            print(f'The sheet \'{f2Sheet}\' doesn\'t exist in {path1}')
        else:
            sheets.add(f2Sheet)

    # Compare the values in each cell
    for sheet in sheets:
        print(f'Cell-level comparision for the sheet \'{sheet}\'...')
        s1 = f1[sheet]
        s2 = f2[sheet]

        for r in s1.index:
            for c in s1.columns:
                v1 = s1.at[r, c]
                v2 = s2.at[r, c]
                if (v1 == v2) or (pd.isna(v1) and pd.isna(v2)):
                    continue
                else:
                    print(f'[{r}, {c}]: {v1} {v2}')
