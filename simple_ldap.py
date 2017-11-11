import ldap
import sys
import time
import csv
import read_employees
import write_results

server = "ldap://ed-p-gl.emea.nsn-net.net:389"
#server = "ldap://ed-p-gl.emea.nsn-net.net:389"
#dn = 'cn=BOOTMAN_Acc,ou=SystemUsers,ou=Accounts,o=NSN'
dn = 'cn=Nokia_asset_management_system_Acc,ou=SystemUsers,ou=Accounts,o=NSN'
#pw = 'Eq4ZVLXqMbKbD4th'
pw = 'EqnFvvKfAc4p2bau'
base_dn = 'ou=Internal,ou=People,o=NSN'

non_sync = 0

ldap_client = None

MANAGER_LEVEL_MAP = {
    '0': 'N-6',
    '5000': 'N-5',
    '100000': 'N-4',
    '500000': 'N-3',
}

def ldap_init():
    if ldap_client is not None:
        return ldap_client
    l = ldap.initialize(server)
    try: 
        #l.start_tls_s()
        res = l.simple_bind_s(dn, pw)
        print 'result: ', res
    except ldap.INVALID_CREDENTIALS:
        print "Your username or password is incorrect."
        return None
    except ldap.LDAPError, e:
        if type(e.message) == dict and e.message.has_key('desc'):
            print e, e.message['desc']
        else: 
            print e
        return None
    finally:
        return l

ldap_client = ldap_init()

def print_managers_by_uid(lc, uid):
    attrs = ['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity']
    res = lc.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,), attrs)
    print res[0]
    if res[0][1].has_key('nsnManagerAccountName'):
        get_user_by_uid(lc, res[0][1]['nsnManagerAccountName'][0])

def get_management_level(uid):
    #l = ldap_init()
    l = ldap_client
    if l is None:
        return None
    level = 0 # the lowest mngm level is 6
    attrs = ['uid', 'nsnManagerAccountName']
    cur_uid = uid
    while True:
        res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(cur_uid,), attrs)
        if res:
            if res[0][1].has_key('nsnManagerAccountName'):
                level = level + 1
                cur_uid = res[0][1]['nsnManagerAccountName'][0]
            else:
                return level
        else:
            print('failed to get uid from ldap, uid=%s'%(cur_uid))
            return level

def get_manager(uid):
    # get manager info of the user with uid
    #l = ldap_init()
    l = ldap_client
    attrs = ['uid', 'uidNumber', 'nsnManagerAccountName']
    res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,), attrs)
    if res:
        mgr_uid = res[0][1]['nsnManagerAccountName'][0]
    res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(mgr_uid,), attrs)
    if res:
        return res[0][1]

def get_manager_verbose(uid):
    l = ldap_init()
    res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,))
    if res:
        mgr_uid = res[0][1]['nsnManagerAccountName'][0]
    res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(mgr_uid,))
    if res:
        return res[0][1]

def get_user_by_uidnumber(uidnum, attrs=['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street']):
    l = ldap.initialize(server)
    try: 
        #l.start_tls_s()
        res = l.simple_bind_s(dn, pw)
        print 'result: ', res
    except ldap.INVALID_CREDENTIALS:
        print "Your username or password is incorrect."
        sys.exit()
    except ldap.LDAPError, e:
        if type(e.message) == dict and e.message.has_key('desc'):
            print e, e.message['desc']
        else: 
            print e
        sys.exit()
    finally:
        if attrs:
            res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uidNumber=%s'%(uidnum,), attrs)
        else:
            res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uidNumber=%s'%(uidnum,))
            print res
        return res

def get_user_by_part_uid(basestrs, found, attrs):
    exceptional = []
    if basestrs == []:
        return found
    for ex in basestrs:
        for lastchr in range(ord('0'), ord('9')+1) + range(ord('a'), ord('z')+1):
            wildcard = ex + unichr(lastchr)
            #print ex, wildcard
            filt = '(&(uid=%s*) (nsnCity=Hangzhou))'%(wildcard,)
            #print filt
            try:
                res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, filt, attrs)
            except ldap.LDAPError, e:
                print e
                exceptional.append(wildcard)
                continue
            for _, u in res:
                try:
                    if 'Xincheng' in u['street'][0]:
                        found.append(u)
                        print 'xincheng user: ', u
                except:
                    pass
    if exceptional != []:
        print('still we have exceptional', exceptional)
        found = get_user_by_part_uid(exceptional, found, attrs)
                
    return found

def is_present(uidnum):
    return 'YES'

def preprocessing(users):
    updated = []
    for user in users:
        for k, v in user.iteritems():
            user[k] = v[0]
        updated.append(user)
    return updated

def get_managers_locally(uid, users):
    lm = 'NA'
    gm = 'NA'
    for user in users:
        if uid == user['uid']:
            try:
                lm = user['nsnManagerAccountName']
            except:
                pass
            break
    for user in users:
        if lm == user['uid']:
            try:
                gm = user['nsnManagerAccountName']
            except:
                pass
            break
    return lm, gm

def get_managers(uid, users):
    n_5 = ''
    n_4 = ''
    n_3 = ''
    for user in users:
        if uid == user['uid']:
            

def save_user_info(users):
    fields = ['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street',
              'displayName', 'employeeType', 'nsnTeamName', 'Present', 'lineManager',
              'groupManager', 'nsnApprovalLimit']
    users = preprocessing(users)
    userinfo = []
    teams = dict()
    with open('ttemployees.csv', 'w') as csvfd:
        writer = csv.DictWriter(csvfd, fieldnames=fields)
        writer.writeheader()
        for user in users:
            user['Present'] = is_present(user['uidNumber'])
            user['lineManager'] = 'anonymous' #get_manager(user['uid'])['uid'][0]
            user['groupManager'] = 'anonymous' #get_manager(user['lineManager'])['uid'][0]
            user['lineManager'], user['groupManager'] = get_managers_locally(user['uid'], users)
            writer.writerow(user)
            userinfo.append(user)
            teams.setdefault(user['nsnTeamName'], [])
            teams[user['nsnTeamName']].append(user)
    write_results.write(fields, userinfo, 'employee sheet')
    save_teams(teams)
    save_user_group(users)

def save_user_group(users):
    n2grps = {}
    n3grps = {}
    n4grps = {}
    n5grps = {}
    print len(users)
    for user in users:
        if user['nsnApprovalLimit'] == '0':
            n5grps.setdefault(user['nsnManagerAccountName'], [])
            n5grps[user['nsnManagerAccountName']].append(user)
        elif user['nsnApprovalLimit'] == '5000':
            n4grps.setdefault(user['nsnManagerAccountName'], [])
            n4grps[user['nsnManagerAccountName']].append(user)
        elif user['nsnApprovalLimit'] == '100000':
            n3grps.setdefault(user['nsnManagerAccountName'], [])
            n3grps[user['nsnManagerAccountName']].append(user)
        elif user['nsnApprovalLimit'] == '500000':
            n2grps.setdefault(user['nsnManagerAccountName'], [])
            n2grps[user['nsnManagerAccountName']].append(user)
        else:
            print 'not expect such user:', user
    fields = ['n-3 manager', 'n-4 manager', 'n-5 manager', 'name', 'present']
    results = []
    for n3 in n3grps:
        for n4 in n3grps[n3]:
            newu = dict()
            newu['n-3 manager'] = n3
            newu['n-4 manager'] = n4['displayName']
            newu['n-5 manager'] = 'na'
            newu['name'] = n4['displayName']
            newu['present'] = n4['Present']
            results.append(newu)
            if not n4grps.has_key(n4['uid']):
                continue
            for n5 in n4grps[n4['uid']]:
                newu = dict()
                newu['n-3 manager'] = n3
                newu['n-4 manager'] = n4['displayName']
                newu['n-5 manager'] = n5['displayName']
                newu['name'] = n5['displayName']
                newu['present'] = n5['Present']
                results.append(newu)
                if not n5grps.has_key(n5['uid']):
                    continue
                for u in n5grps[n5['uid']]:
                    newu = dict()
                    newu['n-3 manager'] = n3
                    newu['n-4 manager'] = n4['displayName']
                    newu['n-5 manager'] = n5['displayName']
                    newu['name'] = u['displayName']
                    newu['present'] = u['Present']
                    results.append(newu)
    hz_n4 = []
    for k, v in n3grps.iteritems():
        hz_n4 += v
    print 'hz n4:', hz_n4
    for mngr, members in n4grps.iteritems():
        if mngr not in [hz4['uid'] for hz4 in hz_n4]:
            print 'non-hz n4', mngr
            for n5 in members:
                newu = dict()
                newu['n-3 manager'] = 'N/A'
                newu['n-4 manager'] = mngr
                newu['n-5 manager'] = n5['displayName']
                newu['name'] = n5['displayName']
                newu['present'] = n5['Present']
                results.append(newu)
                if not n5grps.has_key(n5['uid']):
                    continue
                for u in n5grps[n5['uid']]:
                    newu = dict()
                    newu['n-3 manager'] = 'N/A'
                    newu['n-4 manager'] = mngr
                    newu['n-5 manager'] = n5['displayName']
                    newu['name'] = u['displayName']
                    newu['present'] = u['Present']
                    results.append(newu)
    print len(results)
    write_results.write(fields, results, 'group sheet')
    print n3grps.keys()
    print n2grps
    
def save_teams(teams):
    teaminfo = []
    for k, v in teams.iteritems():
        mngr = v[0]
        t = dict()
        for u in v:
            if u['nsnApprovalLimit'] > mngr['nsnApprovalLimit']:
                mngr = u
        t['teamname'] = k
        t['manager'] = mngr['displayName']
        t['managerlevel'] = MANAGER_LEVEL_MAP[mngr['nsnApprovalLimit']]
        t['total'] = len(v)
        t['actual'] = len([u for u in v if u['Present']])
        #print k, 'total: ', len(v), ' managed by ', mngr['displayName']
        teaminfo.append(t)
    fields = ['teamname', 'manager', 'managerlevel', 'total', 'actual']
    write_results.write(fields, teaminfo, 'team sheet')

def compare_users(users):
    nid_diff = []
    hr_nsn_ids = read_employees.read_employees()
    ldap_nsn_ids = [u['uidNumber'] for u in users]
    print('ldap nsn ids: ', len(ldap_nsn_ids))
    for hr_nsn_id in hr_nsn_ids:
        if hr_nsn_id not in ldap_nsn_ids:
            nid_diff.append(hr_nsn_id)
            get_user_by_uidnumber(hr_nsn_id)
    #print('nsnid gap:', len(nid_diff))
    return nid_diff

if __name__ == '__main__':
        '''
        l = ldap.initialize(server)
        try: 
            #l.start_tls_s()
            res = l.simple_bind_s(dn, pw)
            print 'result: ', res
        except ldap.INVALID_CREDENTIALS:
            print "Your username or password is incorrect."
            sys.exit()
        except ldap.LDAPError, e:
            if type(e.message) == dict and e.message.has_key('desc'):
                print e, e.message['desc']
            else: 
                print e
            sys.exit()
        finally:
        '''
        l = ldap_init()
        if l:
            filt = '(&(uid=%s*) (nsnCity=hangzhou))'%('ab')
            attrs = ['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street', 'displayName', 'employeeType',
                     'nsnTeamName', 'nsnApprovalLimit']
            #attrs = ['+']
            if non_sync:
                rids = []
                for fstch in range(ord('a'), ord('z')+1):
                    for sndch in range(ord('0'), ord('9')+1) + range(ord('a'), ord('z')+1):
                        wildcard = unichr(fstch) + unichr(sndch)
                        filt = '(&(uid=%s*) (nsnCity=hangzhou))'%(wildcard,)
                        rid = l.search_ext(base_dn, ldap.SCOPE_SUBTREE, filt, attrs, timeout=20)
                        rids.append((wildcard, rid))
                users = []
                exceptions = []
                polls = 0
                resps = []
                while polls < 100:
                    if len(resps) + len(exceptions) == len(rids):
                        print("all queries DONE!!!")
                        break
                    print len(resps), len(rids)

                    for wildcard, rid in rids:
                        if wildcard in resps:
                            #print("already done %s"%(wildcard))
                            continue
                        try:
                            res = l.result(rid, True, 50)
                        except ldap.LDAPError, e:
                            print e, wildcard
                            exceptions.append(wildcard)
                            continue
                        #print res
                        if res[0] is not None:
                            resps.append(wildcard)
                            if res[1] != []:
                                for _, u in res[1]:
                                    if 'Xincheng' in u['street'][0]:
                                        users.append(u)
                        else:
                            print("None res for %s"%(wildcard))
                    polls = polls + 1
                print len(users)
                # second round
                rids = []
                for ex in exceptions:
                    for ch in range(ord('0'), ord('9')+1) + range(ord('a'), ord('z')+1):
                        wildcard = ex + unichr(ch)
                        filt = '(&(uid=%s*) (nsnCity=hangzhou))'%(wildcard,)
                        rid = l.search_ext(base_dn, ldap.SCOPE_SUBTREE, filt, attrs, timeout=20)
                        rids.append((wildcard, rid))
                polls = 0
                resps = []
                exceptions = []
                while polls < 100:
                    if len(resps) + len(exceptions) == len(rids):
                        print("all queries DONE!!!")
                        break
                    #print len(resps), len(rids)

                    for wildcard, rid in rids:
                        if wildcard in resps:
                            #print("already done %s"%(wildcard))
                            continue
                        try:
                            res = l.result(rid, True, 50)
                        except ldap.LDAPError, e:
                            print e, wildcard
                            exceptions.append(wildcard)
                            continue
                        else:
                            pass
                            #print wildcard, 'has been done!'
                        #print res
                        if res[0] is not None:
                            resps.append(wildcard)
                            if res[1] != []:
                                for _, u in res[1]:
                                    if 'Xincheng' in u['street'][0]:
                                        users.append(u)
                        else:
                            print("None res for %s"%(wildcard))
                    polls = polls + 1
                print('total number users: %d'%(len(users)))
        
                '''
                poll_num = 0
                raw_res = (None, None)
                print rid
                users = []
                while raw_res[0] is None and poll_num < 300:
                    print 'still polling, %d'%(poll_num)
                    raw_res = l.result(rid, 1)
                    poll_num = poll_num + 1
                    time.sleep(1)
                res_num = 0
                print raw_res[1]
                users.append(raw_res[1][0])
                while res_num < 1000:
                    raw_res = l.result(rid, 0)
                    if raw_res[0]:
                        users.append(raw_res[1][0])
                    else:
                        print 'empty result...'   
                    res_num = res_num + 1
                for _, u in users:
                    print u['sn'][0], u['uid'][0], u['uidNumber'][0]
                '''
            else:
                exceptional = []
                users = []
                for fstch in range(ord('a'), ord('z')+1):
                    #users = get_user_by_part_uid(unichr(fstch), users, attrs)
                    #continue
                    for sndch in range(ord('0'), ord('9')+1) + range(ord('a'), ord('z')+1):
                        wildcard = unichr(fstch) + unichr(sndch)
                        filt = '(&(uid=%s*) (nsnCity=hangzhou))'%(wildcard,)
                        #print filt
                        try:
                            res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, filt, attrs)
                        except:
                            exceptional.append(wildcard)
                            continue
                        #users = ldaphelper.get_search_results(res)
                        for _, u in res:
                            try:
                                if 'Xincheng' in u['street'][0]:
                                    users.append(u)
                                    #print 'xincheng user: ', u
                                if u['nsnStatus'][0].upper() != 'ACTIVE':
                                    print('invalid status, ', u)
                                #print u['nsnTeamId'][0], u['nsnBusinessGroupShortName'][0], u['sn'][0], u['uid'][0], u['uidNumber'][0]
                            except:
                                pass
                        if res:
                            #print res[0][1].keys()
                            pass
                print('Hi, the following filters need further processing...')
                print(exceptional)
                # further processing for those exceptional
                exceptional_wth_three_chr = []
                for ex in exceptional:
                    for thdch in range(ord('0'), ord('9')+1) + range(ord('a'), ord('z')+1):
                        wildcard = ex + unichr(thdch)
                        filt = '(&(uid=%s*) (nsnCity=Hangzhou))'%(wildcard,)
                        try:
                            res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, filt, attrs)
                        except ldap.LDAPError, e:
                            print e, wildcard
                            exceptional_wth_three_chr.append(wildcard)
                            continue
                        for _, u in res:
                            try:
                                if 'Xincheng' in u['street'][0]:
                                    users.append(u)
                                    #print 'xincheng user: ', u
                            except:
                                pass
                        #if res != []: print res[0][1].keys()
                if exceptional_wth_three_chr != []:
                    print('still we have exceptional', exceptional_wth_three_chr)
                    users = get_user_by_part_uid(exceptional_wth_three_chr, users, attrs)
                            
                print('total number of users located in HZ: %d'%(len(users)))
                #get_us er_by_uid(l, 'haitchen')a
                save_user_info(users)
                #print(compare_users(users))
