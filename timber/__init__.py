#! /usr/bin/python3
# -*- coding: utf-8 -*-


import logging as timber

timber.basicConfig(level=timber.INFO,
                   format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                   datefmt='%a, %d %b %Y %H:%M:%S')
