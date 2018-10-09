#!/usr/bin/python
# -*- coding: utf-8 -*-

"""

@author: Jiannan Hua
last edited: 
"""

#################################################
import os
import sys
import email
from imapclient import IMAPClient

hostname = 'mail.ustc.edu.cn'
username = 'USERNAME@mail.ustc.edu.cn'
password = None
dest_dir = r'E:\PATH\download' # 文件保存地址，要求存在
#file_dir = r'E:\TA\homework\down.txt'
existing_file_list = os.listdir(dest_dir)
print(existing_file_list)
cover_same_name_file = False

if password==None:
    password = input('Input password:')


imap = IMAPClient(hostname, ssl=True)
try:
    imap.login(username, password)
except imap.Error:
    print('Logon failed')
    os.system('pause')
    sys.exit(1)
# print(imap)

imap.select_folder('hw', readonly=True) #第一个参数是邮件所在的文件夹，收件箱对应‘Inbox’
result = imap.search('UNSEEN')
print("Total email number:", len(result))

msgdict = imap.fetch(result, ['BODY.PEEK[]'])

#f = open(file_dir, 'r')
#saved = []
#for line in f:
#    if line[-1] == '\n':
#        saved.append(line[:-1])
#    else:
#        saved.append(line[:])
#f.close()
#print(saved)

#f = open(file_dir, 'w')

for message_id, message in msgdict.items():
    print()
    e = email.message_from_string(message[b'BODY[]'].decode('utf-8'))
    
    subject = email.header.make_header(email.header.decode_header(e['SUBJECT']))
    mail_from = email.header.make_header(email.header.decode_header(e['From']))

    print("Subject:",subject)
    print("From:",mail_from)
    
    #if str(subject) in saved:
    #    continue
    #else:
        #print(subject, file=f)
    #    saved.append(str(subject))
    
    for part in e.walk():
        if part.is_multipart():
            continue
        if part.get('Content-Disposition') is None:
            continue #跳过没有附件部分

        file_data = part.get_payload(decode=True)

        file_name = part.get_filename()
        dename = email.header.decode_header(file_name)[0] # 函数返回一个只有一项的list

        # dename=(编码后的文件名，编码方式)
        if dename[1]==None: #不用解码
            name = dename[0]
        else: #解码并存为str
            name = str(dename[0], dename[1]) 

        if name in existing_file_list and cover_same_name_file == False:
            print("File",name, "has existed")
            continue

        print("downloading", name)
        name_b = name.encode('utf-8') #转码为bytes
     
        os.chdir(dest_dir)
        try:
            with open(name_b, 'wb') as fp:
                fp.write(file_data)
                existing_file_list.append(name)
        except:
            print("Download failed")

#print(saved)
#for i in saved:
#    print(i, file=f)

#f.close()

imap.logout()