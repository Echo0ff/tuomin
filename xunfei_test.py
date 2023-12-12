#!/usr/bin/python
# -*- coding: UTF-8 -*-
import time
import urllib.request
import urllib.parse
import json
import hashlib
import base64
#接口地址
url ="http://ltpapi.xfyun.cn/v1/ner"
#开放平台应用ID
x_appid = "b99a0bff"
#开放平台应用接口秘钥
api_key = "7a95266c999537c42f98759e1fbf7535"
#语言文本
TEXT="根据《中华人民共和国审计法》第十六条、第二十二条和第二十三条的规定，连云港市审计局派出审计组，自2017年12月6日"


def main():
    body = urllib.parse.urlencode({'text': TEXT}).encode('utf-8')
    param = {"type": "dependent"}
    x_param = base64.b64encode(json.dumps(param).replace(' ', '').encode('utf-8'))
    x_time = str(int(time.time()))
    x_checksum = hashlib.md5(api_key.encode('utf-8') + str(x_time).encode('utf-8') + x_param).hexdigest()
    x_header = {'X-Appid': x_appid,
                'X-CurTime': x_time,
                'X-Param': x_param,
                'X-CheckSum': x_checksum}
    req = urllib.request.Request(url, body, x_header)
    result = urllib.request.urlopen(req)
    result = result.read()
    print(result.decode('utf-8'))
    return


if __name__ == '__main__':
    main()