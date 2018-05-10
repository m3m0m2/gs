#!/usr/bin/python

from ss import *


def main():
  a = SSArea(3)
  
  a.loadCsv('a.csv')

  print('a.csv')
  print(a)

  a[0][0] = 5
  a[1][1] = '"'
  a[2][1] = ""
  a[2][2] = None

  a.writeCsv('b.csv')

  print('b.csv')
  print(a)
  


if __name__ == '__main__':
    main()
