# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 16:04:09 2020

@author: Cesar
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select
from copy import copy, deepcopy

import time
import openpyxl
import openpyxl.worksheet.cell_range
from openpyxl.styles import Alignment

