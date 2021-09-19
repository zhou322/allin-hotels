import email
import imaplib
import re
from email.header import decode_header
from html import unescape
import csv
import os

download_path = 'local/'

mailbox = '"INBOX.nuwegexclusief reservation"'

host = ''
post = ''
user = ''
password = ''

BOOKER = "Gegevens boeker:"
HOTEL = "Gegevens hotel:"
BOOKING = "Boekingsgegevens:"
COST = "Kostenspecificatie:"


def html_to_plain_text(html):
    text = re.sub('<head.*?>.*?</head>', ' ', html, flags=re.M | re.S | re.I)
    text = re.sub(r'<a\s.*?>', ' HYPERLINK ', text, flags=re.M | re.S | re.I)
    text = re.sub('<.*?>', ' ', text, flags=re.M | re.S)
    text = re.sub(r'(\s*\n)+', '\n', text, flags=re.M | re.S)
    return unescape(text)


def start_mailbox(host, post, uesr, passwrod):
    try:
        conn = imaplib.IMAP4_SSL(host, post)
        conn.login(uesr, passwrod)
        print("Connect to {0}:{1} successfully".format(host, post))
        return conn
    except BaseException as e:
        print("Connect to {0}:{1} failed".format(host, post), e)


def get_body(email):
    if email.is_multipart():
        return get_body(email.get_payload(0))
    else:
        return email.get_payload(decode=True)


def search(conn):
    r, d = conn.search(None, 'FROM "info@nuwegexclusief.nl" SUBJECT "Hotel Confirmation"')
    return r, d


def get_info_by_key(body, key, count):
    test = body.decode()
    text = html_to_plain_text(test)
    lines = text.split('\n')
    # for l in lines:
    #     print(l)
    start = lines.index(key) + 1
    stop = start + count
    return lines[start:stop]


def extract_info(line):
    result = line.split(":")
    if len(result) == 1:
        return result
    else:
        return result[1]


def get_info_by_index(body, index, count):
    test = body.decode()
    text = html_to_plain_text(test)
    lines = text.split('\n')
    stop = index + count
    initial_lines = []
    for l in lines[index:stop]:
        l1 = re.sub(r'\t', '', l)
        l2 = re.sub(r'\xa0', '', l1)
        initial_lines.append(l2.strip())

    return initial_lines


def is_valid_subject(subject):
    return re.match(r"^Hotel Confirmation:", subject)


def is_valid_sender(sender):
    return sender == 'info@nuwegexclusief.nl' or '<info@nuwegexclusief.nl>'


def covert_to_rows(email_id_list):
    initial_list = []
    for email_id in email_id_list:
        row = convert_email_csv_row(email_id)
        print(row)
        initial_list.append(row)
    return initial_list


def convert_email_csv_row(email_id):
    res, msg = connection.fetch(email_id, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode(encoding)
            # decode email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)

            if is_valid_subject(subject) and is_valid_sender(From):
                body = get_body(msg)
                booker_name, booker_address_1, booker_address_2, booker_country, booker_tel, booker_email_address = get_info_by_index(
                    body, 5, 6)
                hotel_name, hotel_address_1, hotel_address_2, hotel_country, hotel_tel, hotel_email_address = get_info_by_index(
                    body, 12, 6)
                boekingsnummer, klantnummer, aantal_volwassenen, aantal_kinderen, s1, s2, datum_aankomst, datum_vertrek, kamers = map(
                    extract_info, get_info_by_index(body, 18, 9))

                order_number = subject.split(' ')
                return order_number[
                           2], booker_name, booker_address_1, booker_address_2, booker_country, booker_tel, booker_email_address, hotel_name, hotel_address_1, hotel_address_2, hotel_country, hotel_tel, hotel_email_address, boekingsnummer, klantnummer, aantal_volwassenen, aantal_kinderen, s1, s2, datum_aankomst, datum_vertrek, kamers
            else:
                print('invalid subject', subject)
                print('invalid From', From)


def write_to_csv(rows, delete_file=True):
    filename = os.path.join(os.path.dirname(__file__), "local/hotel.csv")
    if delete_file and os.path.isfile(filename):
        os.remove(filename)

    data_file = open(filename, 'x')
    csv_writer = csv.writer(data_file, quoting=csv.QUOTE_NONNUMERIC)
    count = 0

    for r in rows:
        if count == 0:
            header = ['order_number', 'booker_name', 'booker_address_1', 'booker_address_2', 'booker_country',
                      'booker_tel',
                      'booker_email_address', 'hotel_name', 'hotel_address_1', 'hotel_address_2', 'hotel_country',
                      'hotel_tel', 'hotel_email_address', 'boekingsnummer', 'klantnummer', 'aantal_volwassenen',
                      'aantal_kinderen', 's1', 's2', 'datum_aankomst', 'datum_vertrek', 'kamers']
            csv_writer.writerow(header)
            count += 1
        if r is not None:
            csv_writer.writerow(r)
    data_file.close()


connection = start_mailbox(host, post, user, password)
connection.select(mailbox)
result, data = search(connection)
ids = data[0]
email_id_list = ids.split()
rows = covert_to_rows(email_id_list)

write_to_csv(rows)

connection.close()
connection.logout()
