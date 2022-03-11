import openpyxl
import pandas as pd
import bs4 as bs
import requests
import codecs
import lxml

def asktckr():
    print("Welcome to Steve's DCF generator")
    asktckr.var = input("What is the ticker of the stock you would like to run a valuation on?  ")

    return

def verifytckr():
    print(asktckr.var)

    return

def executescript():
    asktckr()
    verifytckr()
    return

executescript()