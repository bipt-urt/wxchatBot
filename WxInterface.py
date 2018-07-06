import WxBot

class WxInterface:
	def __init__(self):
		self.bot = WxBot.WxBot()
		self.wxToken = {}
	
	def getToken(self):
		return self.wxToken
	
	def login(self, QRCodeFilename="qrcode.jpg", autoOpen=True):
		import os, time, json, sys
		self.wxToken["loginQRToken"] = self.bot.getLoginToken()
		try:
			with open(QRCodeFilename, "wb") as QRCode:
				QRCode.write(self.bot.getQRCode(self.wxToken["loginQRToken"]))
		except:
			print("写入二维码错误")
			sys.exit()
		print("使用手机扫描二维码图片" + str(QRCodeFilename) + "以登录")
		if autoOpen:
			os.system('call %s' % QRCodeFilename)
		
		while True:
			wxLoginStatus = self.bot.getLoginStatus(self.wxToken["loginQRToken"])
			if wxLoginStatus[1] == "201":
				print("已经扫描二维码，请在手机上点击登录")
			elif wxLoginStatus[1] == "200":
				print("请稍候...")
				self.wxToken["redirectURL"] = self.bot.getWxRedirectURL(wxLoginStatus[0])
				break
			time.sleep(1)
		
		wxWebWechatToken = self.bot.wxRedirect(self.wxToken["redirectURL"])
		self.wxToken["skey"] = wxWebWechatToken.split("<skey>")[1].split("</skey>")[0]
		self.wxToken["wxsid"] = wxWebWechatToken.split("<wxsid>")[1].split("</wxsid>")[0]
		self.wxToken["wxuin"] = wxWebWechatToken.split("<wxuin>")[1].split("</wxuin>")[0]
		self.wxToken["pass_ticket"] = wxWebWechatToken.split("<pass_ticket>")[1].split("</pass_ticket>")[0]
		if self.wxToken["pass_ticket"] == "":
			print("无法获取用户凭据，可能当天登陆次数过多，触发登录频率限制")
			sys.exit()
		self.wxToken.pop("redirectURL")
		
		self.wxInitData = json.loads(self.bot.wxInit(self.wxToken))
		self.wxToken["displayname"] = self.wxInitData["User"]["NickName"]
		self.wxToken["username"] = self.wxInitData["User"]["UserName"]
		
		self.wxContacts = json.loads(self.bot.getContact(self.wxToken))
		
		print("==========你好，" + self.wxToken["displayname"] + "！==========")
	
	def printRecentContacts(self):
		if len(self.wxInitData["ContactList"]):
			print("最近联系人为:")
			for recentCommunicatePerson in self.wxInitData["ContactList"]:
				displayName = recentCommunicatePerson["RemarkName"] or recentCommunicatePerson["NickName"]
				print("\t" + displayName)
			print("\t目前共有" + str(self.wxContacts["MemberCount"]) + "位好友")
		else:
			print("没有最近联系人")
	
	def findFriend(self, keyword):
		findResult = []
		for contacts in self.wxContacts["MemberList"]:
			if contacts["NickName"].find(keyword) != -1 or contacts["RemarkName"].find(keyword) != -1:
				findResult.append({'NickName': contacts["NickName"], 'RemarkName': contacts["RemarkName"], 'UserName': contacts["UserName"]})
		return findResult
	
	def sendMessageUI(self):
		print('发送消息')
		while True:
			contactTo = input("请输入要发送消息的联系人:")
			if contactTo == "":
				return
			searchedList = self.findFriend(contactTo)
			if len(searchedList):
				print("找到以下联系人")
				break
			else:
				print("\t未找到任何联系人")
		for person in searchedList:
			print("\t\t*" + person["NickName"] + " " + person["RemarkName"] + " " + person["UserName"])
		if len(searchedList) > 1:
			contactToId = input("请输入需要联系的人员或群ID:")
		else:
			contactToId = searchedList[0]["UserName"]
		
		wxMessage = input("请输入消息:\n")
		if wxMessage != "":
			self.bot.sendMessage(self.wxToken, contactToId, wxMessage)
	
	def exportContactUI(self):
		print("正在尝试导出联系人列表")
		self.bot.exportContect(self.wxContacts)
		print("已导出联系人列表到friendslist.xlsx")
	
	def locationStaticUI(self):
		print("正在针对联系人地域信息进行统计：\n", end="")
		locationStatic = {}
		for person in self.wxContacts["MemberList"]:
			personLocation = person["Province"]#+person["City"]
			if personLocation == "":
				continue
			if personLocation not in locationStatic:
				locationStatic[personLocation] = 1
			else:
				locationStatic[personLocation] += 1
		locationStaticTmp = sorted(locationStatic.items(), key=lambda x: x[1], reverse=True)
		locationStatic = []
		for element in locationStaticTmp[:10]:
			locationStatic.append(element)
		self.bot.exportECharts(locationStatic)
		print("已完成")
	
	def exportGroupUI(self):
		print('导出群组信息')
		self.bot.exportGroup(self.wxContacts)