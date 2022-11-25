import argparse
import csv
import codecs
from datetime import datetime

import os



description_account = {
    'Crédit STRIPE PAYMENTS UK LTD': 2020,
    "E-banking Ordre à Monika Feniello et Noémie Favre": 2020,
    "Crédit Innova Versicherungen AG": 4070,
    "Crédit Twist Lab SARL": 6002,
    "Crédit VILLE DE NYON": 6002,
    "Crédit COMMUNE DE LA TOUR-DE-PEILZ": 6002,
    "Système de recouvrement direct Viseca Payment Services SA": 2041,
    "E-banking Ordre à Twist Lab": 4100,
    "E-banking Ordre à VAUDOISE GENERALE": 4070,
    "E-banking Ordre à Caisse AVS de la fédération patrona": 4070,
    "E-banking Ordre à Fonds interprofessionnel de prévoya": 4070,
    "E-banking Ordre permanent à Gregory Favre": 4000,
    "E-banking Ordre (eBill) M-Budget Mobile": 4920,
    "E-banking Ordre à Administration fédérale des contrib": 7300,
    "Cotisation annuelle carte de crédit N° 547988XXXXXX6543": 7200,
    "E-banking Ordre à Chemins de fer fédéraux suisses CFF": 4800,
}
description_cleanup = {
    'Crédit STRIPE PAYMENTS UK LTD': 'Stripe - mon arbre enchanté',
    "E-banking Ordre à Monika Feniello et Noémie Favre": "Remboursement mon arbre enchanté",
    "Système de recouvrement direct Viseca Payment Services SA": "Paiement CC",
    "E-banking Ordre (eBill) M-Budget Mobile": "Abonnement mobile",
    "E-banking Ordre à Chemins de fer fédéraux suisses CFF": "Abonnement CFF",
}


def transform_raiffeisen(filepath):
    f = codecs.open(filepath, encoding='latin-1')
    reader = csv.reader(f, delimiter=';')
    # Date, account, description, incoming amount, outgoing amount
    for (iban, t_date, description, amount, balance, v_date) in reader:
        if iban == 'IBAN':
            # header
            continue
        if amount == '0':
            continue
        date = datetime.strptime(t_date, '%Y-%m-%d %H:%M:%S.0').strftime('%d.%m.%Y')
        amount = float(amount)

        if amount > 0:
            incoming = '{:.2f}'.format(amount)
            outgoing = ''
        else:
            incoming = ''
            outgoing = '{:.2f}'.format(abs(amount))
        account = str(description_account.get(description, ''))
        description = description_cleanup.get(description, description)
        print('\t'.join((date, account, description, incoming, outgoing)))


def is_valid_file(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        return arg


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Transform Raiffeisen CSV file into tabulated data')
    parser.add_argument('filepath', metavar='FILE', type=lambda x: is_valid_file(parser, x),
                        help='path to the excel file')
    args = parser.parse_args()
    transform_raiffeisen(args.filepath)
