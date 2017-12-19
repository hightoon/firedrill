#--*-- encoding=utf-8 --*--

import ldap
import sys
import time
import csv
import read_employees
import write_results
import RegistrationInfo
from readreg import read_regs


server = "ldap://ed-p-gl.emea.nsn-net.net:389"
dn = 'cn=Nokia_asset_management_system_Acc,ou=SystemUsers,ou=Accounts,o=NSN'
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

TT_SITE_CODE = '72011003'

manager_info_cache = []
LOCAL_TOPS = [] # the top node of management tree
MNGR_EXPECTED = {}
MNGR_STATS = {}
MNGRS_WITH_TT_BOSS = []
REG_USERS = None
REG_EMPLOYEES = None
REG_EXTERN = None
MATCHED = {}

def get_manager_cache(uid):
    for m in manager_info_cache:
        if uid == m['uid']:
            return m
    return None

def put_manager_cache(mngr):
    global manager_info_cache
    mngr['islm'] = True
    manager_info_cache.append(mngr)

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

try:
    ldap_client = ldap_init()
except Exception as e:
    print 'ldap init failed: ', e

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
    attrs = ['uid', 'uidNumber', 'nsnManagerAccountName', 'displayName', 'nsnSiteCode']
    res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,), attrs)
    if res:
        try:
            mgr_uid = res[0][1]['nsnManagerAccountName'][0]
        except KeyError:
            print uid, res[0][1]
            mgr_uid = 'm14wang' # if no manager, set self as mike wang
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

def get_user_by_uid(uid, attrs=['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street', 'displayName', 'nsnSiteCode', 'nsnApprovalLimit']):
    l = ldap_init()
    if attrs:
        res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,), attrs)
    else:
        res = l.search_s(base_dn, ldap.SCOPE_SUBTREE, 'uid=%s'%(uid,))
    if res:
        res = res[0][1]
        for k, v in res.iteritems():
            res[k] = v[0]
        return res


def get_user_by_uidnumber(uidnum, attrs=['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street', 'nsnSiteCode']):
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
    global MATCHED
    for u in REG_EMPLOYEES:
        if uidnum == u.person.uidnum and '诺基亚通信创新软件园' not in u.location:
            MATCHED[u.nickname] = u.person.uidnum
            return 'YES'
    return 'NO'

def preprocessing(users):
    updated = []
    for user in users:
        for k, v in user.iteritems():
            user[k] = v[0]
        updated.append(user)
    return updated

def get_managers_locally(uid, users):
    lm = ''
    gm = ''
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
                lm = get_manager(user['uid'])
                gm = lm['nsnManagerAccountName']
            break
    return lm, gm

def get_managers(uid, users):
    user = None
    lm = None
    gm = None
    for u in users:
        if u['uid'] == uid:
            user = u
            break
    if user is None:
        user = get_manager_cache(uid)
        if user is None:
            user = get_user_by_uid(uid)
            put_manager_cache(user)
        lm = get_manager_cache(user['nsnManagerAccountName'])
        if lm is None:
            lm = get_user_by_uid(user['nsnManagerAccountName'])
            put_manager_cache(lm)
        gm = get_manager_cache(lm['nsnManagerAccountName'])
        if gm is None:
            gm = get_user_by_uid(lm['nsnManagerAccountName'])
            put_manager_cache(gm)
        return user, lm, gm
    for u in users:
        if u['uid'] == user['nsnManagerAccountName']:
            lm = u
            break
    if lm is None:
        # not located in hz
        lm = get_manager_cache(user['nsnManagerAccountName'])
        if lm is None:
            lm = get_user_by_uid(user['nsnManagerAccountName'])
            put_manager_cache(lm)
        gm = get_manager_cache(lm['nsnManagerAccountName'])
        if gm is None:
            gm = get_user_by_uid(lm['nsnManagerAccountName'])
            put_manager_cache(gm)
    else:
        for u in users:
            if u['uid'] == lm['nsnManagerAccountName']:
                gm = u
            else:
                gm = get_manager_cache(lm['nsnManagerAccountName'])
                if gm is None:
                    gm = get_user_by_uid(lm['nsnManagerAccountName'])
                    put_manager_cache(gm)
    return user, lm, gm

def save_user_info(users):
    fields = ['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street',
              'displayName', 'employeeType', 'nsnTeamName', 'Present', 'lineManager',
              'groupManager', 'nsnApprovalLimit', 'nsnSiteCode']
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
    #write_results.write(fields, userinfo, 'employee sheet')
    #save_teams(teams)
    #save_user_as_groups(users)
    sorted_users, maxlv = generate_management_tree(userinfo)
    # check if top manager with TT boss must be done before removing top mgr
    get_top_managers_with_tt_boss(sorted_users)
    remove_unique_top_mgr(sorted_users)
    generate_mngr_stats(sorted_users)
    fields = ['manager-%d'%(l) for l in range(1, maxlv+1)]
    fields.append('Present')
    write_results.write(fields, sorted_users, 'management tree sheet',
                    expected=MNGR_EXPECTED, stats=MNGR_STATS, ttboss=MNGRS_WITH_TT_BOSS,
                    external=REG_EXTERN)

def generate_management_tree(users):
    max_levels = 0
    for user in users:
        mngrs = []
        mngr = user
        levels = 1
        while mngr['nsnSiteCode'] == TT_SITE_CODE:
            mngrs.append(mngr)
            next_mngr = get_manager_cache(mngr['nsnManagerAccountName'])
            if next_mngr == None:
                next_mngr = get_user_by_uid(mngr['nsnManagerAccountName'])
                next_mngr['Present'] = is_present(next_mngr['uidNumber'])
                put_manager_cache(next_mngr)
            levels = levels + 1
            mngr = next_mngr
        mngrs.append(mngr)
        mngrs.reverse()
        user['managers'] = mngrs
        if levels > max_levels:
            max_levels = levels
    print 'max levels: ', max_levels
    for user in users:
        for i in range(1, max_levels+1):
            if i - 1 < len(user['managers']):
                user['manager-%d'%(i)] = user['managers'][i-1]['displayName']
            else:
                user['manager-%d'%(i)] = ''
    new_users = sorted(users, key=lambda k: k['managers'], reverse=True)
    return new_users, max_levels

def generate_mngr_stats(users):
    ''' get manager based fire-drill statistics from management tree
        input:
            users - must be users after calling generate_management_tree
    '''
    for user in users:
        for i, m in enumerate(user['managers'][:-1]):
            if i == 0 and m['displayName'] not in LOCAL_TOPS:
                MNGR_EXPECTED.setdefault(m['displayName'], 0) # top manager always NOT in TT, so she's not expected
            else:
                MNGR_EXPECTED.setdefault(m['displayName'], 1)
            MNGR_EXPECTED[m['displayName']] += 1
        if user['Present'] == 'YES':
            for m in user['managers'][:-1]:
                if not m.has_key('Present'):
                    print '%s not required to join fire  drill'%(m['displayName'])
                    print m
                    MNGR_STATS.setdefault(m['displayName'], 0)
                else:
                    if m['Present'] == 'YES':
                        MNGR_STATS.setdefault(m['displayName'], 1) # check participatation of herself
                    else:
                        MNGR_STATS.setdefault(m['displayName'], 0)
                        print 'manager: %s not present'%(m['displayName'])
                MNGR_STATS[m['displayName']] += 1

def get_top_managers_with_tt_boss(userlist):
    global MNGRS_WITH_TT_BOSS
    tops = set([(u['managers'][0]['displayName'], u['managers'][0]['nsnManagerAccountName']) for u in userlist])
    for n, mn in tops:
        for mgr in manager_info_cache:
            if mn == mgr['uid'] and mgr['nsnSiteCode'] == TT_SITE_CODE:
                MNGRS_WITH_TT_BOSS.append(n)
    print MNGRS_WITH_TT_BOSS

def remove_unique_top_mgr(userlist):
    global LOCAL_TOPS
    mgr_list_set = set(['+'.join(map(lambda x: x['displayName'], u['managers'][:2])) for u in userlist])
    mgr_counter = {}
    for m in mgr_list_set:
        mgr_counter.setdefault(m.split('+')[0], 0)
        mgr_counter[m.split('+')[0]] += 1
    uniq = []
    for k, v in mgr_counter.iteritems():
        if v == 1:
            uniq.append(k)
    for user in userlist:
        if (user['managers'][0]['displayName'] in uniq 
            and user['managers'][0]['displayName'] not in MNGRS_WITH_TT_BOSS):
            if len(user['managers']) == 2:
                user['managers'][0] = user['managers'][1] # duplicate self
            else:
                user['managers'] = user['managers'][1:]
            for i in range(6):
                user['manager-%d'%(i+1,)] = ''
            for i, m in enumerate(user['managers']):
                user['manager-%d'%(i+1,)] = m['displayName']
            if user['manager-1'] not in LOCAL_TOPS:
                LOCAL_TOPS.append(user['manager-1'])

# not used
def save_user_as_groups(users):
    n2grps = {}
    n3grps = {}
    n4grps = {}
    n5grps = {}
    print len(users)
    for user in users:
        if user['nsnApprovalLimit'] != '':
            lm, n4m, n3m = get_managers(user['nsnManagerAccountName'], users)
            if lm['nsnApprovalLimit'] == '5000' or lm['nsnApprovalLimit'] == '0': #manager can be 0 or 5000
                n5grps.setdefault(lm['displayName'], [])
                if not is_existing_user(user['displayName'], n5grps[lm['displayName']]):
                    n5grps[lm['displayName']].append(user)
                if n4m['nsnApprovalLimit'] == '500000':
                    n3grps.setdefault(n4m['displayName'], [])
                    if not is_existing_user(lm['displayName'], n3grps[n4m['displayName']]):
                        n3grps[n4m['displayName']].append(lm)
                elif n4m['nsnApprovalLimit'] == '100000' or n4m['nsnApprovalLimit'] == '5000':
                    n4grps.setdefault(n4m['displayName'], [])
                    if not is_existing_user(lm['displayName'], n4grps[n4m['displayName']]):
                        n4grps[n4m['displayName']].append(lm)
                    n3grps.setdefault(n3m['displayName'], [])
                    if not is_existing_user(n4m['displayName'], n3grps[n3m['displayName']]):
                        n3grps[n3m['displayName']].append(n4m)
                else:
                    print 'n5 %s reports to %s'%(lm['displayName'], n4m['displayName'])
            elif lm['nsnApprovalLimit'] == '100000':
                n4grps.setdefault(lm['displayName'], [])
                if not is_existing_user(user['displayName'], n4grps[lm['displayName']]):
                    n4grps[lm['displayName']].append(user)
                n3grps.setdefault(n4m['displayName'], [])
                if not is_existing_user(lm['displayName'], n3grps[n4m['displayName']]):
                    n3grps[n4m['displayName']].append(lm)
            elif lm['nsnApprovalLimit'] == '500000':
                n3grps.setdefault(lm['displayName'], [])
                if not is_existing_user(lm['displayName'], n3grps[lm['displayName']]):
                    n3grps[lm['displayName']].append(lm)
            else:
                n3grps.setdefault(user['displayName'], [])
                print '%s directly reports to %s, geeee!'%(user['displayName'], lm['displayName'])
    employees = []
    total = 0
    for n3, members in n3grps.iteritems():
        print n3, 'has following N-4s'
        print [m['displayName'] for m in members]
        for n4 in members:
            if n4['nsnSiteCode'] == TT_SITE_CODE:
                total = total + 1
            if not n4grps.has_key(n4['displayName']):
                user = dict()
                user['name'] = n4['displayName']
                user['n-3'] = n3
                user['n-4'] = ''
                user['n-5'] = ''
                employees.append(user)
                continue
            for n5 in n4grps[n4['displayName']]:
                if n5['nsnSiteCode'] == TT_SITE_CODE:
                    total = total + 1
                if not n5grps.has_key(n5['displayName']):
                    user = dict()
                    user['name'] = n5['displayName']
                    user['n-3'] = n3
                    user['n-4'] = n4['displayName']
                    user['n-5'] = ''
                    employees.append(user)
                    continue
                for u in n5grps[n5['displayName']]:
                    user = dict()
                    user['name'] = u['displayName']
                    user['n-3'] = n3
                    user['n-4'] = n4['displayName']
                    user['n-5'] = n5['displayName']
                    employees.append(user)
                    if u['nsnSiteCode'] == TT_SITE_CODE:
                        total = total + 1
    print 'total: %d, employees: %d'%(total, len(employees))
    print 'done, start writing groups'
    validate_output(employees, users)
    fields = ['n-3', 'n-4', 'n-5', 'name']
    write_results.write(fields, employees, 'groups')

def is_existing_user(dn, users):
    for user in users:
        if user.has_key('displayName') and dn == user['displayName']:
            return True
        if user.has_key('name') and dn == user['name']:
            return True
    return False

def validate_output(out, allusers):
    for u in allusers:
        if not is_existing_user(u['displayName'], out):
            print u

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
    print n4grps.keys()
    print n3grps.keys()
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
        mngr = None
        t = dict()
        for u in v:
            for o in v:
                if o['nsnManagerAccountName'] == u['uid']:
                    mngr = u
        t['teamname'] = k
        if mngr:
            t['manager'] = mngr['displayName']
            t['managerlevel'] = MANAGER_LEVEL_MAP[mngr['nsnApprovalLimit']]
        else:
            t['manager'] = 'nobody'
            t['managerlevel'] = 'na'
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
    #global REG_USERS
    #global REG_EMPLOYEES
    #global REG_EXTERN
    REG_USERS = read_regs()  # read all registration data from local csv file (default Form_data.csv)
    if REG_USERS:
        REG_EMPLOYEES = [u for u in REG_USERS if u.employment=='nsb']
        REG_EXTERN = [u for u in REG_USERS if u.employment=='ext']
    l = ldap_init()
    if l is None:
        raise SystemExit
    if l:
        filt = '(&(uid=%s*) (nsnCity=hangzhou))'%('ab')
        attrs = ['uid', 'uidNumber', 'nsnManagerAccountName', 'nsnCity', 'street', 'displayName', 'employeeType',
                 'nsnTeamName', 'nsnApprovalLimit', 'nsnSiteCode']
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
                                #if 'Xincheng' in u['street'][0]:
                                if u['nsnSiteCode'][0] == TT_SITE_CODE:
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
                                #if 'Xincheng' in u['street'][0]:
                                if u['nsnSiteCode'] == TT_SITE_CODE:
                                    users.append(u)
                    else:
                        print("None res for %s"%(wildcard))
                polls = polls + 1
            print('total number users: %d'%(len(users)))
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
                        except:
                            pass
                    if res:
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
