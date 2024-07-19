# -*- coding: utf-8 -*-
import os
import sys

CWD = os.path.dirname(os.path.realpath(__file__))
ROOT_DIR = os.path.dirname(CWD)
sys.path.append(ROOT_DIR)

from zk import ZK


conn = None
zk = ZK('192.168.1.31', port=4370)
try:
    conn = zk.connect()
    # Get all users (will return list of User object)
    # users = conn.get_users()
    # for u in users:
    #     print(u)
    template = conn.get_user_template(uid=722, temp_id=5)  # temp_id is the finger to read 0~9
    # Get all fingers from DB (will return a list of Finger objects)
    fingers = conn.get_templates()
    # for fin in fingers:
    #     print(fin)
    #     print(type(fin.json_pack()['template']))
    #     print(len(fin.json_pack()['template']))
    #     print((fin.json_pack()['template']))
    # print(len(template.json_pack('template')))
    print(type(template))
    print(template.dump())
    print(type(template.json_pack()['template']))
    # print(len(template.json_pack()['template']))
    print((template.json_pack()['template']))

except Exception as e:
    print ("Process terminate : {}".format(e))
finally:
    if conn:
        conn.disconnect()