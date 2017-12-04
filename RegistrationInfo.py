'''
    RegistrationInfo class
'''

class RegistrationInfo(object):
    def __init__(self, employment, regtime, wechat_nickname=None, uidnum=None, company=None):
        self.employment = employment
        if employment == 'nsb':
            self.person = Employee(uidnum)
        if employment == 'ext':
            self.person = Visitor(company)
        self.nickname = wechat_nickname
        self.regtime = regtime

    def __str__(self):
        if self.employment == 'nsb':
            return 'employee ID: %s, Wechat NickName: %s'%(self.person.uidnum, self.nickname)
        else:
            return 'company: %s, wechat nickname: %s'%(self.person.company, self.nickname)


class Employee(object):
    def __init__(self, uidnum):
        if uidnum is not None:
            self.uidnum = uidnum


class Visitor(object):
    def __init__(self, company):
        if company is not None:
            self.company = company
