from exchangelib import DELEGATE, Account, Credentials, \
    EWSDate, EWSDateTime, EWSTimeZone, Configuration, NTLM, CalendarItem, Message, \
    Mailbox, Attendee, Q, Build, Version
from exchangelib.folders import Calendar

import subprocess # useful for passwordeval

import configparser
import html2text
import datetime
import os
import sys

def main():
    if len(sys.argv) > 1:
        config_file_path = sys.argv[1]
    else:
        config_file_path = os.path.join(
            os.path.dirname(os.path.realpath(__file__)),
            'exchange-calendar-to-org.cfg')

    config = configparser.ConfigParser(converters={'list': lambda x: [i.strip() for i in x.split(',')]})
    config.read(config_file_path)

    settings = config['Settings']

    email = settings.get('email')
    username = settings.get('username', email)

    server_url = settings.get('server_url', None)
    auth_type = settings.get('auth_type', None)
    version = settings.getlist('server_version', None)

    passwordeval = settings.get('passwordeval', None)

    if passwordeval:
        password = eval(passwordeval)
    else:
        password = settings.get('password')

    sync_days = int(settings.get('sync_days'))
    org_file_path = settings.get('org_file')
    tz_string = settings.get('timezone_string')

    calendar_names = settings.getlist("calendar_names",None)
    orgpreamble = settings.get("orgpreamble", None)

    tz = EWSTimeZone.timezone(tz_string)

    credentials = Credentials(username=username, password=password)

    if server_url is None:
        account = Account(
            primary_smtp_address=email,
            credentials=credentials,
            autodiscover=True,
            access_type=DELEGATE)
    else:
        if version:
            version = list(map(int,version))
            version = Version(build=Build(*version))
        if auth_type == "NTLM":
            auth_type=NTLM

        server = Configuration(
            server=server_url,
            credentials=credentials,
            version=version,
            auth_type=auth_type)
        account = Account(
            primary_smtp_address=email,
            config=server,
            autodiscover=False,
            access_type=DELEGATE)

    now = datetime.datetime.now()
    end = now + datetime.timedelta(days=sync_days)

    text = []
    if orgpreamble:
        text.append("".join(map(lambda x: "#"+ x, orgpreamble.splitlines(True)))+"\n\n")
    text.append('* Exchange calendars\n')

    calendars = [account.calendar]
    if calendar_names:
        calendars = calendars + list(map(lambda x: account.calendar // x, calendar_names))

    for cal in calendars:
        items = cal.filter(
            start__lt=tz.localize(EWSDateTime(end.year, end.month, end.day)),
            end__gt=tz.localize(EWSDateTime(now.year, now.month, now.day)), )


        text.append('** ' + cal.name)
        text.append('\n')
        for item in items:
            text.append(get_item_text(item, tz))
            text.append('\n')

    f = open(org_file_path, 'w')
    f.write(''.join(text))


def get_item_text(item, tz):
    text = []
    text.append('*** ' + item.subject)
    start_date = item.start
    start_date_text = get_org_date(start_date, tz)

    end_date = item.end
    end_date_text = get_org_date(end_date,tz)
    text.append('<' + start_date_text + '>--<' + end_date_text + '>')
    if item.location is not None:
        text.append('Location: ' + item.location)
    if item.required_attendees is not None or item.optional_attendees is not None:
        text.append('Attendees:')

    if item.required_attendees is not None:
        for person in item.required_attendees:
            text.append('- ' + str(person.mailbox.name))

    if item.optional_attendees is not None:
        for person in item.optional_attendees:
            text.append('- ' + str(person.mailbox.name))

    if item.body is not None:
        text.append('')
        text.append('**** Information')
        text.append(html2text.html2text(item.body).replace('\n\n', '\n'))

    return '\n'.join(text)


def get_org_date(date, tz):
    if type(date) is EWSDate:
       return date.strftime('%Y-%m-%d %a')
    else:
        return date.astimezone(tz).strftime('%Y-%m-%d %a %H:%M')


if __name__ == '__main__':
    main()
