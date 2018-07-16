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
		try:
			return urllib.request.urlopen(url).read().decode(codecs)
		except:
			print("网络无法连接")
	
	def getRaw(self, url):
		'Send a HTTP GET request and dont decode it in order to get raw data'
		try:
			return urllib.request.urlopen(url).read()
		except:
			print("网络无法连接")
	
	def post(self, url, postData, codecs="utf-8"):
		'Send a HTTP POST request to target server'
		try:
			return urllib.request.urlopen(url, json.dumps(postData, ensure_ascii=False).encode(codecs)).read().decode(codecs)
		except:
			print("网络无法连接")
	
	def postPayload(self, url, payload, codecs="utf-8"):
		'Send a HTTP POST request to a target server via raw payload'
		try:
			return urllib.request.urlopen(url, payload).read().decode(codecs)
		except:
			print("网络无法连接 while postPayload")
	
	def options(self, url, codecs="utf-8"):
		'Send a HTTP OPTIONS request to a target server with (possible) url argument'
		requestHandle = urllib.request.Request(url)
		requestHandle.get_method = lambda: 'OPTIONS'
		try:
			return urllib.request.urlopen(requestHandle).read().decode(codecs)
		except:
			print("网络无法连接")
	
	def getCookie(self):
		return self.cj