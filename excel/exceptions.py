#!/usr/bin/env python3
# -*- coding: utf-8 -*-

class ExcellException(Exception):
    def __init__(self, msg = ''):
        super(self.__class__, self).__init__(msg)
        self.msg = msg
    
    def print_msg(self):
        print(self.msg)

class ReadError(ExcellException):
    def __init__(self, filename):
        self.msg = '[%s] 파일을 읽는중에 오류가 발생하였습니다.' % (filename)
