import simple_ldap

def test_get_user_by_uidnum(uidnum, attrs):
    res = simple_ldap.get_user_by_uidnumber(uidnum, attrs=attrs)
    print res

def test_get_user_by_uid(uid):
    print simple_ldap.get_user_by_uid(uid)

def test_get_manager(uid):
    print simple_ldap.get_manager(uid)

def test_get_manager_verbose(uid):
    print simple_ldap.get_manager_verbose(uid)

def test_get_management_level(uid):
    print simple_ldap.get_management_level(uid)

def test_remove_unique_top_mgr(userlist):
    simple_ldap.remove_unique_top_mgr(userlist)
    print userlist

test_data = [u'61418924', u'10209924', u'10187674', u'10187675', u'10193194', u'10201434', u'10187676', u'10187683', u'10187632', u'10193196', u'10175545', u'61258125', u'10215329', u'10212606', u'61336267', u'61372494', u'61374041', u'61376860', u'61403800', u'61421099', u'62062816', u'62063023', u'62063807', u'62069330', u'62069626', u'62069868', u'62074617', u'62074691', u'62074869', u'62078522', u'62078560', u'62079448', u'62079575', u'62081091', u'62089400', u'62089411', u'62092241', u'62092264', u'62092525', u'62094704', u'62097425', u'62099564', u'62099635', u'62100198', u'62100689']

if __name__ == '__main__':
    #for data in test_data:
    #    test_get_user_by_uidnum(data)
    try:
        test_get_manager('haitchen')
        test_get_manager_verbose('haitchen')
        test_get_management_level('haitchen')
        test_get_management_level('guxu')
        test_get_management_level('rayzheng')
        test_get_user_by_uidnum('61465429', None)
        test_get_user_by_uidnum('61278396', None)
        test_get_user_by_uidnum('61402704', None)
        test_get_user_by_uidnum('10163277', None)
        test_get_user_by_uid('diding')
        test_get_user_by_uid('haitchen')
        test_get_user_by_uid('sberrahi')
        test_get_user_by_uid('dawchen')
        test_get_user_by_uid('xuechen')
    except:
        pass
    test_remove_unique_top_mgr(
        [{'managers': [{'displayName': 'haitong'}, {'displayName': 'zhang qi'}], 'manager-1': 'haitong', 'manager-2': 'zhang qi', 'manager-3':'', 'manager-4': '', 'manager-5': '', 'manager-6': ''},
         {'managers': [{'displayName': 'haitong'}, {'displayName': 'zhang yulong'}], 'manager-1': 'haitong', 'manager-2': 'zhang yulong', 'manager-3': '', 'manager-4': '', 'manager-5': '', 'manager-6': ''},
         {'managers': [{'displayName': 'mike wang'}, {'displayName': 'zhang qi'}, {'displayName': 'zhang xin'}], 'manager-1': 'mike wang', 'manager-2': 'zhang qi', 'manager-3': 'zhang xin', 'manager-4': '', 'manager-5': '', 'manager-6': ''},
         {'managers': [{'displayName': 'mike wang'}, {'displayName': 'zhang qi'}], 'manager-1': 'mike wang', 'manager-2': 'zhang qi', 'manager-3': '', 'manager-4': '', 'manager-5': '', 'manager-6': ''}
        ]
    )
