#!/usr/bin/python3

from __future__ import print_function
import json
import os
import csv
import codecs
import zipfile
import httplib2

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Limmud 2019 Web App'


def transliterate(word):
    translit_table={'\u0410': 'A', '\u0430': 'a',
        '\u0411': 'B', '\u0431': 'b',
        '\u0412': 'V', '\u0432': 'v',
        '\u0413': 'G', '\u0433': 'g',
        '\u0414': 'D', '\u0434': 'd',
        '\u0415': 'E', '\u0435': 'e',
        '\u0416': 'Zh', '\u0436': 'zh',
        '\u0417': 'Z', '\u0437': 'z',
        '\u0418': 'I', '\u0438': 'i',
        '\u0419': 'I', '\u0439': 'i',
        '\u041a': 'K', '\u043a': 'k',
        '\u041b': 'L', '\u043b': 'l',
        '\u041c': 'M', '\u043c': 'm',
        '\u041d': 'N', '\u043d': 'n',
        '\u041e': 'O', '\u043e': 'o',
        '\u041f': 'P', '\u043f': 'p',
        '\u0420': 'R', '\u0440': 'r',
        '\u0421': 'S', '\u0441': 's',
        '\u0422': 'T', '\u0442': 't',
        '\u0423': 'U', '\u0443': 'u',
        '\u0424': 'F', '\u0444': 'f',
        '\u0425': 'Kh', '\u0445': 'kh',
        '\u0426': 'Ts', '\u0446': 'ts',
        '\u0427': 'Ch', '\u0447': 'ch',
        '\u0428': 'Sh', '\u0448': 'sh',
        '\u0429': 'Shch', '\u0449': 'shch',
        '\u042a': '"', '\u044a': '',
        '\u042b': 'Y', '\u044b': 'y',
        '\u042c': "'", '\u044c': "",
        '\u042d': 'E', '\u044d': 'e',
        '\u042e': 'Yu', '\u044e': 'yu',
        '\u042f': 'Ya', '\u044f': 'ya'}
    
    converted_word = ''
    for char in word:
        transchar = ''
        if char in translit_table:
            transchar = translit_table[char]
        else:
            transchar = char
        converted_word += transchar
    return converted_word


def parse_name(name):
    """
    @param name - name in 'Last First' or 'Last.First' or 'Last First$First Last' format
    @return ['Last First', 'First Last']
    """
    n = name.split('.')
    if len(n) == 2:
        return [name.replace('.', ' '), n[1] + ' ' + n[0], transliterate(n[1]), transliterate(n[0])]

    n = name.split('$')
    if len(n) == 2:
        return [n[0], n[1], transliterate(n[1]), '']

    n = name.split(' ')
    if len(n) == 2:
        return [name, n[1] + ' ' + n[0], transliterate(n[1]), transliterate(n[0])]

    return [name, name, transliterate(name), transliterate(name)]


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    # some global variables
    event = 'limmud2019'

    tracks = {}
    tracks_id = 0

    # internal arrays - index == id
    locations = [ '' ]
    presentors = [ '' ]
    categories = [ '' ]

    # colors = [ '#BF532E', '#3DC8C3', '#1485CC', '#E8782C', '#B062A1', '#9EFF56', '#FFF22A', '#1C9ECC', '#FF3C11', '#B7CDFF', '#FF3046', '#5972B0', '#E0331F', '#1CCCC5', '#84CC12' ]

    colors = [ '#EEE',
               '#80C5CA', # dark cyan
               '#B7CDFF', # blue
               '#CCAA92', # brown
               '#9EDF7D', # green
               '#F3827F', # red
               '#CEF2EC', # cyan
               '#E673FF', # magenta
               '#EEE',    # grey
               '#FFFD67', # yellow
               '#FFBC57', # dark brown
               '#DDFF55', # bright green
               '#BF532E', '#3DC8C3', '#1485CC', '#E8782C', '#B062A1', '#9EFF56', '#FFF22A', '#1C9ECC', '#FF3C11', '#B7CDFF', '#FF3046', '#5972B0', '#E0331F', '#1CCCC5', '#84CC12' ]

    # external arrays - as should be written to json files
    microlocations = []
    speakers = []
    sessions = []
    tracks = []

    # read data from Google Sheets
    print('Connect to Google Sheets...')
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1zHDlydoFTHNiMlUE_EUVHGVGVhoEFIIAryVrRldEQ2I'

    # parse the locations
    print('Reading locations...')
    rangeName = 'Locations!A2:E'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('Can not read locations data from Google Sheets')
        return

    for row in values:
      try:
        if row[0]:
            hotel = row[0]
        if row[1]:
            room = row[1]
        else:
            room = ''
        if row[2]:
            hotel_he = row[2]
        if row[3]:
            room_he = row[3]
        else:
            room_he = room
        if len(row) > 4 and row[4]:
            color = row[4]
        else:
            color = ''

        if not room:
            continue

        if hotel != '-':
            location = hotel + ' - ' + room
            location_he = hotel_he + ' - ' + room_he
        else:
            location = room
            location_he = room_he

        location_id = len(locations)
        microlocations.append({'id' : location_id, 'name' : location, 'name_he' : location_he, 'color' : color})
        locations.append(location)

      except IndexError as e:
        continue

    # parse the colors
    print('Reading colors...')
    rangeName = 'Colors!A2:E'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('Can not read colors data from Google Sheets')
        return

    for row in values:
      try:
        if row[0]:
            category = row[0].capitalize()
        else:
            category = ''
        if row[1]:
            color = row[1]
        else:
            color = '#EEE'
        if row[2]:
            category_he = row[2].capitalize()
        else:
            category_he = ''

        if not category:
            continue

        category_id = len(categories)
        tracks.append({'id' : category_id, 'name' : category, 'name_he' : category_he, 'description' : category, 'sessions' : [], 'color' : color})
        categories.append(category)

      except IndexError as e:
        continue


    # parse the presentors
    print('Reading presentors...')
    rangeName = 'Presentors!A2:G'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('Can not read presentors data from Google Sheets')
        return

    for row in values:
        try:
            if row[0]:
                person = row[0]
            else:
                person = ''
            if row[1]:
                descr = row[1]
            else:
                descr = ''
            if row[2]:
                biography = row[2]
            else:
                biography = ''
            if row[3]:
                photo = row[3]
            else:
                photo = 'avatar.png'
            if len(row) > 4 and row[4]:
                person_he = row[4]
            else:
                person_he = person
            if len(row) > 5 and row[5]:
                descr_he = row[5]
            else:
                descr_he = ''
            if len(row) > 6 and row[6]:
                biography_he = row[6]
            else:
                biography_he = ''

            if not person:
                continue

            person_id = len(presentors)
            person_name = parse_name(person)
            person_name_he = parse_name(person_he)
            speakers.append({'id' : person_id, 'name' : person_name[0], 'name2' : person_name[1], 'name_he' : person_name_he[0], 'name2_he' : person_name_he[1], 'sessions' : [], 'photo' : '/images/speakers/' + photo, 'short_biography' : descr, 'short_biography_he' : descr_he, 'long_biography' : biography, 'long_biography_he' : biography_he, 'first_name': person_name[2], 'last_name': person_name[3]})
            presentors.append(person_name[0])

        except IndexError as e:
            continue


    # parse the schedule
    print('Reading schedule...')
    rangeName = 'Schedule!A2:S'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('Can not read speakers data from Google Sheets')
        return

    row_id = 0
    for row in values:
        row_id += 1
        try:
            if row[0]:
                date = row[0]
            if row[1]:
                start = row[1]
            if row[2]:
                end = row[2]
            else:
                if row[1]:
                    hhmm = start.split(':')
                    hour = int(hhmm[0]) + 1
                    if hour == 24:
                        hour = 0
                    if hour < 10:
                        end = '0' + str(hour) + ':' + hhmm[1]
                    else:
                        end = str(hour) + ':' + hhmm[1]
            if row[3]:
                hotel = row[3]
            if row[4]:
                room = row[4]
            if row[5]:
                people = row[5].split(',')
            else:
                people = []
            if row[6]:
                name = row[6]
            else:
                name = ''
            if row[7]:
                languages = row[7].split(',')
                languages.sort()
                language = ', '.join(languages)
            else:
                languages = []
                language = ''
            if row[8]:
                presentation = row[8]
            if row[9]:
                description = row[9]
            else:
                description = ''
            if row[10]:
                hotel_he = row[10]
            if row[11]:
                room_he = row[11]
            if row[12]:
                people_he = row[12].split(',')
            else:
                people_he = []
            if row[13]:
                name_he = row[13]
            else:
                name_he = ''
            if len(row) > 14 and row[14]:
                languages_he = row[14].split(',')
                languages_he.sort()
                language_he = ', '.join(languages_he)
            else:
                languages_he = []
                language_he = ''
            if len(row) > 15 and row[15]:
                presentation_he = row[15]
            if len(row) > 16 and row[16]:
                description_he = row[16]
            else:
                description_he = ''
            if len(row) > 17 and row[17]:
                shabbat = True
            else:
                shabbat = False
            if len(row) > 18 and row[18]:
                recommend = True
            else:
                recommend = False

            if not name:
                continue

            if hotel == '-' and room == '-':
                continue

            if not people_he:
                people_he = people

            # update locations
            if hotel != '-':
                location = hotel + ' - ' + room
                location_he = hotel_he + ' - ' + room_he
            else:
                location = room
                location_he = room_he
            if not location in locations:
                location_id = len(locations)
                microlocations.append({'id' : location_id, 'name' : location, 'name_he' : location_he, 'color' : ''})
                locations.append(location)

            # update speakers
            for i in range(len(people)):
                person_name = parse_name(people[i].strip())
                if not person_name[0] in presentors:
                    person_name_he = parse_name(people_he[i].strip())
                    person_id = len(presentors)
                    speakers.append({'id' : person_id, 'name' : person_name[0], 'name2' : person_name[1], 'name_he' : person_name_he[0], 'name2_he' : person_name_he[1], 'sessions' : [], 'photo' : '/images/speakers/avatar.png', 'short_biography' : '', 'short_biography_he' : '', 'long_biography' : '', 'long_biography_he' : ''})
                    presentors.append(person_name[0])

            # update end_time if session overlaps to the next day
            offset = '+02:00'
            start_time = date + 'T' + start + ':00' + offset
            if end.startswith('0'):
                yymmdd = date.split('-')
                end_time = yymmdd[0] + '-' + yymmdd[1] + '-' + str(int(yymmdd[2]) + 1) + 'T' + end + ':00' + offset
            else:
                end_time = date + 'T' + end + ':00' + offset

            # update tracks
            if False:
            #if presentation == 'лекция':
                category = presentation.capitalize() + ' (' + language + ')'
                category_he = presentation_he + ' (' + language_he + ')'
                language = ''
                language_he = ''
            else:
                category = presentation.capitalize()
                category_he = presentation_he
            if not category in categories:
                category_id = len(categories)
                tracks.append({'id' : category_id, 'name' : category, 'name_he' : category_he, 'description' : category, 'sessions' : [], 'color' : colors[category_id]})
                categories.append(category)

            # update sessions
            session_id = row_id
            session = {'id' : session_id, 'title' : name, 'title_he' : name_he, 'start_time' : start_time, 'end_time' : end_time, 'microlocation' : {}, 'track' : {}, 'speakers' : [], 'audio' : None, 'long_abstract' : description, 'long_abstract_he' : description_he, 'language' : language, 'language_he' : language_he, 'shabbat' : shabbat, 'recommend' : recommend }
            session['microlocation']['id'] = locations.index(location)
            session['microlocation']['name'] = location
            session['microlocation']['name_he'] = location_he
            location_color = ''
            for m in microlocations:
                if m['name'] == location:
                    location_color = m['color']
            session['microlocation']['color'] = location_color
            session['track']['id'] = categories.index(category)
            session['track']['name'] = category
            session['track']['name_he'] = category_he
            for i in range(len(people)):
                person_name = parse_name(people[i].strip())
                person_name_he = parse_name(people_he[i].strip())
                session['speakers'].append({'id' : presentors.index(person_name[0]), 'name' : person_name[0], 'name_he' : person_name_he[0]})
            sessions.append(session)                                                                         

            # update speaker sessions
            for person in people:
                person_name = parse_name(person.strip())
                for speaker in speakers:
                    if speaker['name'] == person_name[0]:
                        speaker['sessions'].append({'id' : session_id, 'title' : name, 'title_he' : name_he})

            # update track sessions
            for track in tracks:
                if track['name'] == category:
                    track['sessions'].append({'id' : session_id, 'title' : name, 'title_he' : name_he})


            print(date + ' ' + start + '-' + end + ' ', end='')
            try:
                print(name)
            except UnicodeEncodeError as e:
                print('')

        except IndexError as e:
            continue

    # write microlocations file
    with codecs.open(event + '/microlocations', 'w', 'utf-8') as jsonfile:
        json.dump(microlocations, jsonfile, indent=4, separators=(',', ': '), ensure_ascii=False)

    # write speakers file
    with codecs.open(event + '/speakers', 'w', 'utf-8') as jsonfile:
        json.dump(speakers, jsonfile, indent=4, separators=(',', ': '), ensure_ascii=False)

    # write tracks file
    with codecs.open(event + '/tracks', 'w', 'utf-8') as jsonfile:
        json.dump(tracks, jsonfile, indent=4, separators=(',', ': '), ensure_ascii=False)

    # write sessions file
    with codecs.open(event + '/sessions', 'w', 'utf-8') as jsonfile:
        json.dump(sessions, jsonfile, indent=4, separators=(',', ': '), ensure_ascii=False)

    # presenter names (for badges)
    # with codecs.open(event + '/presenters.csv', 'w+', 'utf-8') as file1:
    #     for speaker in sorted(speakers, key = lambda i: i['last_name']):
    #         file1.write(speaker['first_name'] + ',' + speaker['last_name'] + ',презентер\n')
    
    # write ZIP file
    os.chdir(event)
    zipf = zipfile.ZipFile('../' + event + '.zip', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk('./'):
        for dir1 in dirs:
            zipf.write(os.path.join(root, dir1))
        for file in files:
            if not file.endswith('zip'):
                zipf.write(os.path.join(root, file))
    zipf.close()


if __name__ == '__main__':
    main()
