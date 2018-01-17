'''
    RegistrationInfo class
'''
import re

class RegistrationInfo(object):
    def __init__(self, employment, regtime, wechat_nickname=None, uidnum=None, company=None, location=None):
        self.employment = employment
        if employment == 'nsb':
            self.person = Employee(uidnum)
        if employment == 'ext':
            self.person = Visitor(company)
        self.nickname = wechat_nickname
        self.regtime = regtime
        self.location = location

    def __str__(self):
        if self.employment == 'nsb':
            return 'employee ID: %s, Wechat NickName: %s, @%s, @%s'%(self.person.uidnum, self.nickname, 
                self.regtime, self.location)
        else:
            return 'company: %s, wechat nickname: %s, @%s, @%s'%(self.person.company, self.nickname, 
                self.regtime, self.location)

    def is_valid_location(self):
        ''' check if user was registered in required area by reading
            longitude and latitude in location info
        '''
        try:
            lon, lat = map(float, re.findall("\d+\.\d+", self.location))
        except:
            print 'no loc info'
            return False
        #if (lon < 120.172274 and lon > 120.167409) and (lat > 30.188606 and lat < 30.189754):
        if lon < 120.172274 and lon > 120.167409:
            return True 
        return False

    def is_valid_time(self):
        return '2017-12-26 10:00:00' < self.regtime < '2017-12-26 11:00:00'


class Employee(object):
    def __init__(self, uidnum):
        if uidnum is not None:
            self.uidnum = uidnum


class Visitor(object):
    def __init__(self, company):
        if company is not None:
            self.company = company
