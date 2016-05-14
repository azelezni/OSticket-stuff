from ldap3 import Connection
from sys import argv
import json
import pymysql

"""
    Get user information from Active Directory with given username (samaccountname)
    and store retrieved information in database.
"""


def mobile_parse(phone):
    """
       Parse mobile number and return in a predefined format: 052-123-4567
    """
    chars = ['+', '/', ' ', '-']
    mobile = phone
    for c in chars:
        mobile = mobile.replace(c, '')
    if mobile[0:3] == '972':  # replace only first 3 letters (in case there is a 972 in the actual number)
        mobile = mobile[3:]
        if mobile[0] != '0':  # check if there is a 0 after 972 (ie +972-052-123-4567)
            mobile = '0' + mobile
    mobile = mobile[0:3] + '-' + mobile[3:6] + '-' + mobile[6:]
    return mobile

''' check for error with username '''
if len(argv) < 1 or argv[1] == '':
    raise IndexError("no username given")

username = argv[1]
users = {}
user_info = {}

''' set AD base, filter, and search attributes '''
ad_base = ''
ad_filter = ''
ad_attributes = []

''' connect to AD '''
ad_conn = Connection('', user='', password='', auto_bind=True)
ad_conn.search(ad_base, ad_filter, attributes=ad_attributes)
info = json.loads(ad_conn.entries[0].entry_to_json())

''' store user information '''
try:
    user_info['mobile'] = mobile_parse(''.join(info['attributes']['mobile']))
except KeyError:
    pass
finally:
    user_info['id'] = ''.join(info['attributes']['employeeID'])
    user_info['job_description'] = ''.join(info['attributes']['jobDescription'])
    user_info['manager_id'] = ''.join(info['attributes']['mgrID'])

''' change filter and attributes to grab user's manager information '''
ad_attributes = ['mail', 'name']
ad_filter = '(&(objectcategory=person)(objectclass=user)(employeeID=%s))' % user_info['manager_id']
ad_conn.search(ad_base, ad_filter, attributes=ad_attributes)
info = json.loads(ad_conn.entries[0].entry_to_json())

''' store manager information '''
try:
    user_info['manager_name'] = ''.join(info['attributes']['name'])
    user_info['manager_mail'] = ''.join(info['attributes']['mail'])
except KeyError:
    pass

''' connect to database '''
db_host = ''
db_user = ''
db_passwd = ''
db_name = ''
mysql_conn = pymysql.connect(host=db_host, user=db_user, passwd=db_passwd, db=db_name)
mysql_cur = mysql_conn.cursor()

''' create form entry values '''
for value in user_info:
    print(value)
    mysql_cur.execute("SELECT id FROM ost_form_entry "
                      "WHERE object_id = (SELECT user_id FROM ost_user_account WHERE username = '%s')" % username)
    entry_id = mysql_cur.fetchone()[0]
    mysql_cur.execute("SELECT id FROM ost_form_field WHERE name = '%s'" % value)
    field_id = mysql_cur.fetchone()[0]
    mysql_cur.execute("UPDATE ost_form_entry_values SET value='%s' "
                      "WHERE entry_id=%s AND field_id=%s;" % (user_info[value], entry_id, field_id))

mysql_conn.commit()
mysql_cur.close()
mysql_conn.close()
