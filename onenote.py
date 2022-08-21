# -*-coding:utf-8-*-
#!/usr/bin/env python3

import mimetypes
import os
import re
import sqlite3
from collections import namedtuple, OrderedDict
from datetime import datetime, timedelta, timezone
from sys import platform
from urllib.parse import urlencode
from zipfile import ZipFile

import requests
from bs4 import BeautifulSoup
from requests import HTTPError
from requests.auth import AuthBase
import time

CLIENT_ID = 'ddbd41a5-c46d-44f2-85b2-d73bbd7bee7d'
CLIENT_SECRET = 'qqgu6ApNZyUmvYgna2WBwK5'
REDIRECT_URI = 'https://login.live.com/oauth20_desktop.srf'

LINUX_DATA_PATH = os.path.expanduser('~/.wiznote/{}/data')
API_BASE = 'https://www.onenote.com/api/v1.0/me/notes'
AUTH_URL = 'https://login.live.com/oauth20_authorize.srf?' + urlencode({
    'response_type': 'code',
    'client_id': CLIENT_ID,
    'redirect_uri': REDIRECT_URI,
    'scope': 'wl.signin office.onenote_create'
})
OAUTH_URL = 'https://login.live.com/oauth20_token.srf'

Document = namedtuple('Document', ['guid', 'title', 'location', 'name', 'url', 'created'])

NOTEBOOK_DICT = {}
SECTION_DICT = {}

class BearerAuth(AuthBase):
    def __init__(self, token):
        self.token = token

    def __call__(self, r):
        r.headers['Authorization'] = 'Bearer ' + self.token
        return r


def get_data_dir():
    while True:
        if platform == 'linux':
            account = input('Input WizNote account: ')
            data_path = LINUX_DATA_PATH.format(account)
        elif platform == 'win32' or platform == 'darwin':
            data_path = input('Input WizNote dir path of "index.db": ')
        else:
            raise Exception('Unsupported platform')

        index_path = os.path.join(data_path, 'index.db')
        if not os.path.isfile(index_path):
            print('Account data not found!')
            continue

        break

    return data_path, index_path


def get_doc_path(data_path, doc):
    if platform == 'linux' or platform == 'darwin':
        return os.path.join(data_path, 'notes', '{%s}' % doc.guid)

    if platform == 'win32':
        return os.path.join(data_path, doc.location.strip('/'), doc.name)

    raise Exception('Unsupported platform')


def get_token(session):
    print('Sign in via: ')
    print(AUTH_URL)
    print('Copy the url of blank page after signed in,\n'
          'url should starts with "https://login.live.com/oauth20_desktop.srf"')

    while True:
        # url = input('URL: ')

        # match = re.match(r'https://login.live.com/oauth20_desktop\.srf\?code=([\w-]{37})', url)
        # print('debug: just input code!')
        # if not match:
        #     print('Invalid URL!')
        #     continue

        # code = match.group(1)
        code = input('input code like *.*******.*******(about 45 chars, between "code=" and "&lc=":')
        break

    resp = session.post(OAUTH_URL, data={
        'grant_type': 'authorization_code',
        'client_id': CLIENT_ID,
        'client_secret': '',
        'code': code,
        'redirect_uri': REDIRECT_URI,
    }).json()

    return resp['access_token']


def create_notebook(session, name):
    #name = input('Pleas input new notebook name(such as WizNote, can\'t be empty).\nName: ')
    #print('Creating notebook: "%s"' % name)
    print('Creating notebook: "%s"' % name)
    resp = session.post(API_BASE + '/notebooks', json={'name': name})
    resp.raise_for_status()

    return resp.json()['id']


def create_section(session, notebook_id, name):
    print('Creating section: "%s"' % name)
    resp = session.post(API_BASE + '/notebooks/%s/sections' % notebook_id, json={'name': name})
    resp.raise_for_status()

    return resp.json()['id']


def get_documents(count):
    data_path, index_path = get_data_dir()

    result = OrderedDict()
    with sqlite3.connect(index_path) as conn:
        sql = """
        SELECT DOCUMENT_GUID, DOCUMENT_TITLE, DOCUMENT_LOCATION, DOCUMENT_NAME, DOCUMENT_URL, DT_CREATED,
          DOCUMENT_PROTECT, DOCUMENT_ATTACHEMENT_COUNT
        FROM wiz_document
        ORDER BY DOCUMENT_LOCATION;
        """.strip()

        cur = conn.execute(sql)
        count_get = 0
        while True:
            row = cur.fetchone()
            if not row:
                break
            # The last transferred content is no longer recorded
            if count_get >= count + 1:
                guid, title, location, name, url, created, protect, attachment_count = row

                if protect:
                    print('Ignore protected document "%s"' % (location + title))
                    continue

                if attachment_count:
                    print('Ignore %d attachment(s) in "%s"' % (attachment_count, location + title))

                docs = result.get(location)
                if not docs:
                    docs = []
                    result[location] = docs

                doc = Document(guid, title, location, name, url, created)
                docs.append(doc)
            count_get = count_get + 1
    num_total = count_get - count
    return data_path, result, num_total


def clean_html(data, doc):
    def parse_datetime(time_str):
        time = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
        return time.replace(tzinfo=timezone(timedelta(hours=-8))).isoformat()

    soup = BeautifulSoup(data, 'lxml')

    pattern = re.compile(r'^index_files/(.+?)$')
    imgs = soup.find_all('img', src=pattern)

    src_file_names = []
    for img in imgs:
        src = img['src']
        match = pattern.fullmatch(src)
        file_name = match.group(1)

        src_file_names.append(file_name)
        img['src'] = 'name:' + file_name

    head_tag = soup.head
    if not head_tag:
        head_tag = soup.new_tag('head')
        soup.html.insert(0, head_tag)

    title_tag = head_tag.title
    if not title_tag:
        title_tag = soup.new_tag('title')

    # to avoid 'Untitled' in title
    title_tag.string = doc.title
    head_tag.insert(0, title_tag)

    created_tag = soup.new_tag('meta', attrs={'name': 'created', 'content': parse_datetime(doc.created)})
    head_tag.insert(1, created_tag)

    if doc.url:
        url_tag = soup.new_tag('p')
        url_tag.string = 'URL: ' + doc.url
        soup.body.insert(0, url_tag)

    return soup.encode('utf-8'), src_file_names


def upload_doc(session, section_id, data_path, doc):
    doc_path = get_doc_path(data_path, doc)

    print('Processing %s%s (%s)' % (doc.location, doc.title, doc.guid))

    with ZipFile(doc_path) as zip_file:
        html_content, src_file_names = clean_html(zip_file.read('index.html'), doc)

        if len(src_file_names) > 5:
            print('Upload may failed if images more than 5')

        data_send = {
            'Presentation': (None, html_content, mimetypes.guess_type('index.html')[0])
        }

        for name in src_file_names:
            data_send[name] = (None, zip_file.read('index_files/' + name), mimetypes.guess_type(name)[0])

    resp = session.post(API_BASE + '/sections/%s/pages' % section_id, files=data_send)
    resp.raise_for_status()

def get_name(location):
    size = len(location.split('/')) - 2
    notebook_name = location.split('/')[1]
    if size == 1:
        section_name = notebook_name
    elif size == 2:
        section_name = location.split('/')[2]
    elif size >= 3:
        section_name = location.split('/')[2] + '-' + location.split('/')[3]
    
    return notebook_name, section_name

def get_id(notebook_name, section_name, session):

    if notebook_name in NOTEBOOK_DICT:
        #if notebook exists, return the notebook id
        notebook_id = NOTEBOOK_DICT[notebook_name]
        if section_name in SECTION_DICT:
        #if section exists ,return the section id
            section_id = SECTION_DICT[section_name]
        else:
        #if section doesn't exists, create the section and add id to SECTION_DICT
            section_id = create_section(session, notebook_id, section_name)
            SECTION_DICT[section_name] = section_id
    else :
        #if notebook doesn't exists, create the notebook and the section, 
        #and add them to NOTEBOOK_DICT and SECTION_DICT
        notebook_id = create_notebook(session, notebook_name)
        section_id = create_section(session, notebook_id, section_name)
        NOTEBOOK_DICT[notebook_name] = notebook_id
        SECTION_DICT[section_name] = section_id
    
    return notebook_id, section_id


def main():
    #Read the count.txt file to get the location where the last transfer ended
    if os.path.exists('count.txt'):
        with open('count.txt','r') as file:
            lastTime = file.read()
        if not lastTime == "":
            count = int(lastTime.split(',')[0])
            notebook_name = lastTime.split(',')[1]
            notebook_id = lastTime.split(',')[2]
            section_name = lastTime.split(',')[3]
            section_id = lastTime.split(',')[4]
            NOTEBOOK_DICT[notebook_name] = notebook_id
            SECTION_DICT[section_name] = section_id
        else:
            count = 0
    else:
        count = 0
    data_dir, docs,num_total = get_documents(count)
    with requests.session() as session:
        token = get_token(session)
        session.auth = BearerAuth(token)
        start_time = time.time()
        num_finish = 1
        print('A total of %d notes to be transferred' % num_total)
        for location, docs in docs.items():
            notebook_name, section_name = get_name(location)
            notebook_id, section_id = get_id(notebook_name, section_name, session)
            for doc in docs:
                # Store the location of this transfer
                with open('count.txt', 'w') as file:
                    file.write(str(count) + "," + notebook_name + "," + str(notebook_id) + "," + section_name + "," + str(section_id))
                upload_doc(session, section_id, data_dir, doc)
                # percentage of finish and time
                now_time = format(time.time()-start_time, '0.1f')+' s'
                percentage = format(num_finish/num_total * 100, '0.1f')+ '%'
                print('%s is finised, take %s' % (percentage, now_time))
                count = count + 1
                num_finish += 1

if __name__ == '__main__':
    try:
        main()
    except HTTPError as e:
        print(e.response.json())
