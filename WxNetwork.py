import urllib.request
import http.cookiejar
import json

class WxNetwork:
	'WxNetwork is a class to handle network resources'
	def __init__(self):
		self.cj = http.cookiejar.CookieJar()
		self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))
		self.opener.addheaders = [
			('User-agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0'),
			('Host', 'wx.qq.com'),
			('Accept', 'application/json, text/plain, */*'),
			('Accept-Language', 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2'),
			('Referer', 'https://wx.qq.com/'),
			('DNT', '1')
		]
		urllib.request.install_opener(self.opener)
	
	def get(self, url, codecs="utf-8"):
		'Send a HTTP GET request to target server'
		return urllib.request.urlopen(url).read().decode(codecs)
	
	def getRaw(self, url):
		'Send a HTTP GET request and dont decode it in order to get raw data'
		return urllib.request.urlopen(url).read()
	
	def post(self, url, postData, codecs="utf-8"):
		'Send a HTTP POST request to target server'
		return urllib.request.urlopen(url, json.dumps(postData, ensure_ascii=False).encode(codecs)).read().decode(codecs)