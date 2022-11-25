#! /usr/bin/python3
import argparse
from collections import namedtuple
import os

import xlrd

HEADER_ROW = 6
XL_DATEMODE = 0  # 0: 1900-based, 1: 1904-based.
ID_ROW = 'RÉFÉRENCE DE TRANS'
DATE_ROW = 'DATE TRANS.'
AMOUNT_ROW = 'MTT BRUT ORIGIN.'
CURRENCY_ROW = 'DEVISE ORIGINE'
RATE_ROW = 'TAUX DEVISE'
SUBTOTAL_ROW = 'MTT BRUT FACTUR.'
SELLER_ROW = 'NOM FOURNIS.'
CATEGORY_ROW = 'CODE MCC'

CATEGORY_TO_ACCOUNT = {
    0:    '7200',  # viseca
    2741: '4940',  # Elektor (pi mag)
    4812: '4300',  # Fix my phone
    4814: '4910',  # Poste
    4816: '4900',  # Arlo, infomaniak, paypal fastspring
    4899: '4901',  # Youtube, google storage
    5045: '4900',  # Computers
    5099: '4900',  # Durable Goods Not Elsewhere Classified
    5311: '4900',  # digitec, banggood
    5331: '4900',  # Variety Stores - fruugo
    5399: '4900',  # Gonser AG
    5499: '4850',  # Special T
    5651: '4900',  # Merchoid
    5732: '4900',  # totonic, interdiscount
    5734: '4901',  # mailgun, shift, autodesk
    5735: '4901',  # apple icloud
    5815: '4901',  # Youtube
    5817: '4900',  # jetbrains, google service
    5818: '4901',  # apple icloud
    5968: '4940',  # Medium
    5999: '4900',  # Paddle
    7011: '4850',  # Lodging–Hotels, Motels, Resorts–not elsewhere classified
    7372: '4901',  # google cloud, postmark, sentry, github, gandi, druide (et c'est faux!)
    7333: '4810',  # shutterstock, istock
    7399: '4901',  # gsuite
    7523: '4800',  # Parking
    8249: '4080',  # Schools, Trade and Vocational Schools - Udemy

}
SELLER_TO_ACCOUNT = {
    'infomaniak.com': '4901',
    'gandi': '4900',
    'google thegrizzlylabs': '4901',
    'google': '4901',
    'google storage': '4901',
    'google gsuite_beyondth': '4901',
    'druide infor antidote': '4900',
    'driautodesk inc.': '4900',
    'paypal iterate cyberd': '4900',
    'whereby': '4920',
    'adobe': '4900',
    'adobe acropro subs': '4900',
    'matthieudesroches.com': '4080',
}
SELLER_CLEANUP = {
    'jetbrains': 'Jetbrains - logiciel',
    'google thegrizzlylabs': 'Google - Grizzly app',
    'google google storage': 'Google - storage',
    'google youtubepremium': 'Google - YouTube',
    'google gsuite_beyondth': 'Google - mail',
    'mailgun technologies': 'Mailgun',
    'apple.com/bill': 'iTunes - Cloud',
    'google cloud emea limi': 'Google - mail',
    'arlo': 'Arlo - caméras - abo',
    'github': 'Github compte pro',
    'paypal banggood': 'Banggood - ?',
    'postmarkapp.com': 'Postmark - envoi mail',
    'shift yearly plan': 'Shift app',
    'débit cotisation annuelle': 'Cotisation annuelle',
    'sentry': 'Sentry - gestion bugs',
    'infomaniak.com': 'Infomaniak - ?',
    'digitec galaxus (onlin': 'Digitec - ?',
    'paypal distrelecsc': 'Distrelec - ?',
    'paypal merchoid': 'Merchoid - ?',
    'paypal banggoodtec': 'Banggood - ?',
    'www.obdev.at': 'LittleSnitch - logiciel',
    'gandi': 'Gandi - domaines',
    'harvest': 'Harvest - logiciel',
    'aws emea': 'Amazon Web Services',
    'driautodesk inc.': 'Autodesk - logiciel',
    'medium annual': 'Medium - veille technologique',
    'microspot.ch': 'Microspot - ?',
    'paddle.net poedit': 'Paddle - poedit - logiciel',
    'paypal iterate cyberd': 'Cyberduck - logiciel',
    'druide infor antidote': 'Antidote - logiciel',
    'phantomjscloud credits': 'PhantomJSCloud',
    'gonser ag': 'Gonser - ?',
    'elektor international': 'Elektor - Raspberry pi mag',
    'postch post ch ag': 'Poste - timbres',
    'special.t by nestle': 'Special T',
    'fix my phone - réparation': 'Fix my phone - réparation',
    'signmax ab': 'Signmax - étiquette boîte aux lettres',
    'google storage': 'Google - storage',
    'whereby': 'Whereby',
    'adobe': 'Adobe - logiciel',
    'adobe acropro subs': 'Adobe - logiciel',
    'totonic gmbh': 'Totonic - Pi-shop',
    'cff prail gland': 'Parking CFF Gland',
    'parking centre ville': 'Parking centre ville, Nyon',
    'paypal udemy': 'Udemy - formation',
    'paypal fruugo': 'Fruugo - ?',
    'auberge de la réunion': 'Auberge de la Réunion, Coinsins',
}

FEE_SELLER = 'frais de traitement 1.5'


def transform_viseca(filepath):
    Transaction = namedtuple('Transaction',
                             ['date', 'amount', 'currency', 'rate', 'subtotal', 'fee', 'total', 'seller', 'account'])

    book = xlrd.open_workbook(filepath)
    sheet = book.sheet_by_index(0)
    header_row = sheet.row_values(HEADER_ROW)
    transactions = {}

    def get_account(values):
        if values[SELLER_ROW] in SELLER_TO_ACCOUNT:
            return SELLER_TO_ACCOUNT[values[SELLER_ROW]]
        if int(values[CATEGORY_ROW]) in CATEGORY_TO_ACCOUNT:
            return CATEGORY_TO_ACCOUNT[int(values[CATEGORY_ROW])]
        else:
            print(f'Unknown category:{values[CATEGORY_ROW]} for seller:{values[SELLER_ROW]}' )
            return ''

    for i in range(HEADER_ROW + 1, sheet.nrows):
        values = dict(zip(header_row, sheet.row_values(i)))
        if values[AMOUNT_ROW] < 0:
            continue
        transaction_id = values[ID_ROW] or f't-{i}'
        seller = SELLER_CLEANUP.get(values[SELLER_ROW], values[SELLER_ROW])
        if transaction_id in transactions:
            old_transaction = transactions[transaction_id]
            if seller == FEE_SELLER:
                seller = old_transaction.seller
            transaction = old_transaction._replace(
                fee=values[SUBTOTAL_ROW],
                total=old_transaction.subtotal + values[SUBTOTAL_ROW],
                seller=seller
            )
        else:
            transaction = Transaction(
                date=xlrd.xldate_as_datetime(values[DATE_ROW], XL_DATEMODE),
                amount=values[AMOUNT_ROW],
                currency=values[CURRENCY_ROW],
                rate=values[RATE_ROW],
                fee=0,
                subtotal=values[SUBTOTAL_ROW],
                total=values[SUBTOTAL_ROW],
                seller=SELLER_CLEANUP.get(values[SELLER_ROW], values[SELLER_ROW]),
                account=get_account(values),
            )
        transactions[transaction_id] = transaction

    ordered_transactions = sorted(transactions.values(), key=lambda t: t.date)

    print('-' * 80)

    for transaction in ordered_transactions:
        print(
            '\t'.join([
                transaction.date.strftime('%d.%m.%Y'),
                transaction.account,
                transaction.seller,
                '',
                '{:.2f}'.format(transaction.total)
                #str(round(transaction.total, 2))
            ])
        )


def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        return arg


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Transform Viseca xls file into tabulated data')
    parser.add_argument('filepath', metavar='FILE', type=lambda x: is_valid_file(parser, x),
                        help='path to the excel file')
    args = parser.parse_args()
    transform_viseca(args.filepath)
