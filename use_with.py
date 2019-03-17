#!/usr/bin/env python
# -*- encoding: utf-8 -*-
"""
@File    :   use_with.py.py    
@Contact :   Johnd0712@hotmail.com
@License :   (C)Copyright 2017-2018, Krynn.cn

@Modify Time         @Author      @Version    @Desciption
-----------------    ---------    --------    -----------
2019-03-17 16:20      Raistlind    1.0         None
"""


# import lib

def open_file():
    try:
        f = open('./static/test.txt', 'r', encoding='utf-8')
        rest = f.read()
        print(rest)
    except:
        pass
    finally:
        f.close()

    with open('./static/test.txt', 'r', encoding='utf-8') as f:
        rest = f.read()
        print(rest)


if __name__ == '__main__':
    open_file()
