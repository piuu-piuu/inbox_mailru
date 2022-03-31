from imap_tools import MailBox, AND
import os
import time
import pandas as pd


class mailru_attachment_saver():

    def __init__(self):
        self.SERVER = 'imap.mail.ru'

        self.LOGIN = 'who@list.ru'
        self.PWD = 'pwd'
        self.SAVETO = 'attachments'

        # 'Входящие' = INBOX;  subfolders delimiter = '/'
        # FOLDER_PATH = '########@mail.ru'
        # FOLDER_PATH = 'INBOX/алиасы/some_inbox@mail.ru'
        self.FOLDER_PATH = 'INBOX/__test'

        # starting senders' list
        self.addresses = []

    def get_filename(self, message, attachment, timestamp=True):
        month = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
        date = message.date_str.split()
        date = date[3] + '-' + str(month[date[2]]) + '-' + date[1]
        try:
            filename = attachment.filename
        except UnicodeEncodeError:
            filename = message.from_ + \
                "_NON-UNICODABLE_" + str(time.time_ns())
        for char in ['/\:<>|*?']:
            filename = filename.replace(char, '-')
        filename = filename.replace('\r', '-')
        filename = filename.replace('\n', '-')
        if timestamp:
            filename = date+' ' + filename
        return filename

    def check_senders(self):
        ex_data = pd.read_excel('mails.xlsx')
        mails = ex_data['ПОЧТА'].values.tolist()
        with open('no_attachments_sent.txt', 'w', encoding='utf-8') as nonsenders:
            for _mail in mails:
                _mail = str(_mail).lower()
                if _mail != 'nan':
                    if _mail not in self.addresses:
                        nonsenders.write(str(_mail)+'\n')

    def parse_attachments(self):
        with MailBox(self.SERVER).login(self.LOGIN, self.PWD) as self.mailbox:
            # readonly for not marking as seen, set "False" to mark messages as seen while fetching

            self.mailbox.folder.set(self.FOLDER_PATH, readonly=True)
            # fetching ALL, automatically marking everything seen
            # for msg in mailbox.fetch(criteria=AND(seen=False), mark_seen=False):
            # fetching UNSEEN, no automatic marking as seen
            dir_path = os.path.dirname(os.path.realpath(__file__))
            new_path = os.path.join(
                dir_path, self.SAVETO, *self.FOLDER_PATH.split('/'))
            if not os.path.exists(new_path):
                os.makedirs(new_path)

            for msg in self.mailbox.fetch(criteria=AND(seen=False), mark_seen=False):
                for att in msg.attachments:

                    # UTF-8 decoding if possible, otherwise use default name; optionally timestamp added to filename
                    filename = self.get_filename(
                        msg, att, timestamp=False)
                    print('attachment found: ' + filename)
                    # adding senders to a sender list
                    if not msg.from_ in self.addresses:
                        self.addresses.append(msg.from_.lower())

                    new_file = os.path.join(new_path, filename)
                    if not os.path.exists(new_file):
                        with open((new_file), 'wb') as f:
                            f.write(att.payload)
                        meta_file = os.path.join(
                            new_path, filename + '.meta.txt')
                        with open((meta_file), 'w') as f:
                            print(msg.date_str + ' ' + msg.from_.lower())
                            f.write(msg.date_str + ' ' + msg.from_.lower())


if __name__ == '__main__':

    s = mailru_attachment_saver()
    s.parse_attachments()
    print(s.addresses)
