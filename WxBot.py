import WxNetwork
class WxBot:
	'WxBot is a Python application to operate your WeChat use web WeChat API'
	
	def __init__(self):
		self.wxNet = WxNetwork.WxNetwork()
	
	def getR(self):
		import random
		randomTicket = str(random.random())[2:] + '1'
		return randomTicket
	
	def getLoginToken(self):
		jsLoginURL = "https://login.wx.qq.com/jslogin?appid=wx782c26e4c19acffb"
		jsLogin = self.wxNet.get(jsLoginURL)
		return jsLogin.split("\"")[1]
	
	def getQRCode(self, QRCodeToken):
		qrCodeURL = "https://login.weixin.qq.com/qrcode/" + str(QRCodeToken)
		return self.wxNet.getRaw(qrCodeURL)

	def getLoginStatus(self, QRCodeToken):
		loginStatusURL = "https://login.wx.qq.com/cgi-bin/mmwebwx-bin/login?loginicon=true&uuid=" + str(QRCodeToken) + "&tip=0&r=" + self.getR() + "&_=1530583685459"
		loginStatus = self.wxNet.get(loginStatusURL)
		loginStatusNumber = loginStatus.split("window.code=")[1].split(";")[0]
		return [loginStatus, loginStatusNumber]
	
	def getWxRedirectURL(self, rawLoginResponse):
		return rawLoginResponse.split("window.redirect_uri=\"")[1].split("\"")[0] + "&fun=new"
	
	def wxRedirect(self, redirectURL):
		return self.wxNet.get(redirectURL)
	
	def buildBaseRequest(eslf, wxToken):
		return {'Uin': wxToken["wxuin"], 'Sid': wxToken["wxsid"], 'Skey': wxToken["skey"], 'DeviceID': 'e756936914066192'}
	
	def wxInit(self, wxToken):
		initURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxinit?r=" + self.getR() + "&pass_ticket=" + wxToken["pass_ticket"]
		postData = {
			'BaseRequest': self.buildBaseRequest(wxToken)
		}
		return self.wxNet.post(initURL, postData)
	
	def getContact(self, wxToken):
		getContactURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxgetcontact?pass_ticket=" + wxToken["pass_ticket"] + "&r=" + self.getR() + "&seq=0&skey=" + wxToken["skey"]
		return self.wxNet.get(getContactURL)
	
	def sendMessage(self, wxToken, sendTo, sendMessage):
		sendMessageURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxsendmsg?pass_ticket=" + wxToken["pass_ticket"]
		postData = {
			'BaseRequest': self.buildBaseRequest(wxToken),
			'Msg': {
				'ClientMsgId': self.getR(),
				'Content': sendMessage,
				'FromUserName': wxToken["username"],
				'LocalID': self.getR(),
				'ToUserName': sendTo,
				'Type': '1'
			},
			'Scene': 0
		}
		return self.wxNet.post(sendMessageURL, postData)
	
	def exportContect(self, _contactsList, _exportFileName = "friendslist.xlsx"):
		import os
		try:
			import xlsxwriter
		except:
			print("需要pip安装XlsxWrite第三方库")
		workbook = xlsxwriter.Workbook(_exportFileName)
		worksheet = workbook.add_worksheet('Contacts')
		sheetdata = [['姓名', '性别', '城市', '签名', '备注']]
		for person in _contactsList["MemberList"]:
			sheetdata.append([person["NickName"], person["Sex"], person["Province"]+person["City"], person["Signature"], person["RemarkName"]])
		row = 0
		for nickname, sex, city, signature, remarkname in sheetdata:
			cell_format = workbook.add_format()
			cell_format.set_pattern(1)  # This is optional when using a solid fill.
			if row % 2 == 0:
				cell_format.set_bg_color('green')
			else:
				cell_format.set_bg_color('yellow')
			worksheet.write(row, 0, nickname, cell_format)
			worksheet.write(row, 1, sex, cell_format)
			worksheet.write(row, 2, city, cell_format)
			worksheet.write(row, 3, signature, cell_format)
			worksheet.write(row, 4, remarkname, cell_format)
			row += 1
		workbook.close()
		Filename = "friendslist.xlsx"
		os.system('call %s' % Filename)
		return None
	
	def exportECharts(self, _locationStatics, _filename = "echarts.htm"):
		import os, json
		chartsList = []
		for location in _locationStatics:
			chartsList.append({'value': location[1], 'name': location[0]})
		with open(_filename, "w", encoding='utf-8') as f:
			f.write("<!DOCTYPE html>\n")
			f.write("<html>\n")
			f.write("	<head>\n")
			f.write('		<meta charset="utf-8">\n')
			f.write("		<script src=\"https://cdn.bootcss.com/echarts/4.1.0.rc2/echarts.min.js\"></script>\n")
			f.write('		<link href="css/allstyle.css" rel="stylesheet" type="text/css"> ')
			f.write("	</head>\n")
			f.write('	<body><div class="outer"><div class="middle"><div class="inner"><div id="login-main">\n')

			f.write("		<!-- 为 ECharts 准备一个具备大小（宽高）的 DOM -->\n")
			f.write("		<div id=\"main\" style=\"width: 600px;height:400px;\"></div>\n")
			f.write("		<script>\n")
			f.write("			var myChart = echarts.init(document.getElementById('main'));\n")
			f.write("			var option = {\n")
			f.write("			tooltip: {\n")
			f.write("				trigger: 'item',\n")
			f.write('				formatter: "{a} <br/>{b}: {c} ({d}%)"\n')
			f.write("			},\n")
			f.write("			legend: {\n")
			f.write("				orient: 'vertical',\n")
			f.write("				x: 'left',\n")
			f.write("				data:" + json.dumps([ele[0] for ele in _locationStatics], ensure_ascii=False) + "\n")
			f.write("			},\n")
			f.write("			series: [\n")
			f.write("				{\n")
			f.write("					name:'访问来源',\n")
			f.write("					type:'pie',\n")
			f.write("					radius: ['50%', '70%'],\n")
			f.write("					avoidLabelOverlap: false,\n")
			f.write("					label: {\n")
			f.write("						normal: {\n")
			f.write("							show: false,\n")
			f.write("							position: 'center'\n")
			f.write("						},\n")
			f.write("						emphasis: {\n")
			f.write("							show: true,\n")
			f.write("							textStyle: {\n")
			f.write("								fontSize: '30',\n")
			f.write("								fontWeight: 'bold'\n")
			f.write("							}\n")
			f.write("						}\n")
			f.write("					},\n")
			f.write("					labelLine: {\n")
			f.write("						normal: {\n")
			f.write("							show: false\n")
			f.write("						}\n")
			f.write("					},\n")
			f.write("					data: \n")
			f.write(json.dumps(chartsList, ensure_ascii=False))
			f.write("				}\n")
			f.write("			]\n")
			f.write("		};\n")
			f.write("		myChart.setOption(option);\n")
			f.write("		</script>\n")
			f.write("	</div></div></div></div></body>\n")
			f.write("</html>\n")
		os.system('call %s' % "echarts.htm")
		return None
	
	def exportGroup(self, contactsList):
		import os
		with open("group.csv", "w", encoding="gb18030") as groupcsv:
			for group in contactsList["MemberList"]:
				if group["UserName"].find("@@") != -1:
					groupcsv.write(group["NickName"] + "," + group["UserName"] + "\n")
		os.system('call %s' % "group.csv")

		return None
	
	def sendImage(self, wxToken, sendTo, imageLocation):
		import os.path
		import hashlib
		if os.path.isfile(imageLocation) == False:
			print("发送图片不存在")
			return False
		sendImageURL = 'https://file.wx.qq.com/cgi-bin/mmwebwx-bin/webwxuploadmedia?f=json'
		self.wxNet.options(sendImageURL)
		payLoad = []
		randomValue = self.getR()
		hashm = ""
		imageContent = None
		with open(imageLocation, 'rb') as f:
			imageContent = f.read()
			hashm = hashlib.new('md5', imageContent).hexdigest()
		print("要发送" + imageLocation + ", MD5值为:" + hashm)
		for item in self.wxNet.getCookie():
			if item.name == "webwx_data_ticket":
				wxToken["webwx_data_ticket"] = item.value
				break
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="id"')
		payLoad.append("")
		payLoad.append("WU_FILE_1")
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="name"')
		payLoad.append("")
		payLoad.append(imageLocation)
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="type"')
		payLoad.append("")
		payLoad.append("image/jpeg")
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="size"')
		payLoad.append("")
		payLoad.append(str(os.path.getsize(imageLocation)))
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="mediatype"')
		payLoad.append("")
		payLoad.append("pic")
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="uploadmediarequest"')
		payLoad.append("")
		payLoad.append('{"UploadType":2,"BaseRequest":{"Uin":' + wxToken["wxuin"] + ',"Sid":"' + wxToken["wxsid"] + '","Skey":"' + wxToken["skey"] + '","DeviceID":"e878530504072308"},"ClientMediaId":1531445280921,"TotalLen":' + str(os.path.getsize(imageLocation)) + ',"StartPos":0,"DataLen":' + str(os.path.getsize(imageLocation)) + ',"MediaType":4,"FromUserName":"' + wxToken["username"] + '","ToUserName":"' + sendTo + '","FileMd5":"' + hashm + '"}')
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="webwx_data_ticket"')
		payLoad.append("")
		payLoad.append(wxToken["webwx_data_ticket"])
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="pass_ticket"')
		payLoad.append("")
		payLoad.append(wxToken["pass_ticket"])
		payLoad.append("-----------------------------" + randomValue)
		payLoad.append('Content-Disposition: form-data; name="filename"; filename="' + imageLocation + '"')
		payLoad.append('Content-Type: image/jpeg')
		payLoad.append("")
		payLoad.append(imageContent)
		payLoad.append("-----------------------------" + randomValue + "--")
		newPayLoad = []
		for element in payLoad:
			if type(element) == str:
				newPayLoad.append(element.encode("utf-8"))
		newPayLoad = b"\r\n".join(newPayLoad)
		ans = self.wxNet.postPayload(sendImageURL, newPayLoad)
		print(ans)
		return True

def unitTest():
	bot = WxBot()
	print("[UT:回放攻击随机数生成(10000)]")
	for x in range(10000):
		bot.getR()
	print("[UT:获取登录凭据(10)]")
	for x in range(10):
		print("\r\t[" + str(x+1) + "/10]" + bot.getLoginToken(), end="")
	print("\n[UT:获取登录二维码]")
	for x in range(10):
		bot.getQRCode(bot.getLoginToken())
		print("\r\t" + str(x+1) + "/10完成", end="")
	print("\n")

if __name__ == "__main__":
	unitTest()