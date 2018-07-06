import urllib.request
import http.cookiejar
import time
import json
import os
import random

try:
	import xlsxwriter
except:
	print("需要pip安装XlsxWrite第三方库")

wxToken = {}
def getR():
	randomTicket = str(random.random())[2:]+'1'#"-1577634346"
	return randomTicket

def dropHTML(_rawData):
	return None

def wxGetLoginToken():
	jsLoginURL = "https://login.wx.qq.com/jslogin?appid=wx782c26e4c19acffb"
	jsLogin = urllib.request.urlopen(jsLoginURL).read().decode("utf-8")
	return jsLogin.split("\"")[1]

def wxGetQRCode(QRCodeToken):
	qrCodeURL = "https://login.weixin.qq.com/qrcode/" + str(QRCodeToken)
	return urllib.request.urlopen(qrCodeURL).read()

def wxGetLoginStatus(QRCodeToken):
	loginStatusURL = "https://login.wx.qq.com/cgi-bin/mmwebwx-bin/login?loginicon=true&uuid=" + str(QRCodeToken) + "&tip=0&r=-1575532129&_=1530583685459"
	loginStatus = urllib.request.urlopen(loginStatusURL).read().decode("utf-8")
	loginStatusNumber = loginStatus.split("window.code=")[1].split(";")[0]
	return [loginStatus, loginStatusNumber]

def getWxRedirectURL(rawLoginResponse):
	return rawLoginResponse.split("window.redirect_uri=\"")[1].split("\"")[0] + "&fun=new"

def wxRedirect(redirectURL):
	return urllib.request.urlopen(redirectURL).read().decode("utf-8")

def buildBaseRequest(wxToken):
	return {'Uin': wxToken["wxuin"], 'Sid': wxToken["wxsid"], 'Skey': wxToken["skey"], 'DeviceID': 'e756936914066192'}

def wxInit(wxToken):
	initURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxinit?r=" + getR() + "&pass_ticket=" + wxToken["pass_ticket"]
	postData = {
		'BaseRequest': buildBaseRequest(wxToken)
	}
	return urllib.request.urlopen(initURL, json.dumps(postData).encode('utf-8')).read().decode('utf-8')

def wxGetContact(wxToken):
	getContactURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxgetcontact?pass_ticket=" + wxToken["pass_ticket"] + "&r=" + getR() + "&seq=0&skey=" + wxToken["skey"]
	return urllib.request.urlopen(getContactURL).read().decode("utf-8")

def wxSendMsg(wxToken, sendTo, sendMessage):
	sendMessageURL = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxsendmsg?pass_ticket=" + wxToken["pass_ticket"]
	postData = {
		'BaseRequest': buildBaseRequest(wxToken),
		'Msg': {
			'ClientMsgId': getR(),
			'Content': sendMessage,
			'FromUserName': wxToken["username"],
			'LocalID': getR(),
			'ToUserName': sendTo,
			'Type': '1'
		},
		'Scene': 0
	}
	return urllib.request.urlopen(sendMessageURL, json.dumps(postData, ensure_ascii=False).encode('utf-8')).read().decode('utf-8')

def exportContect(_contactsList, _exportFileName = "friendslist.xlsx"):
	workbook = xlsxwriter.Workbook(_exportFileName)
	worksheet = workbook.add_worksheet('Contacts')
	sheetdata = [['姓名', '性别', '城市', '签名', '备注']]
	for person in _contactsList["MemberList"]:
		sheetdata.append([person["NickName"], person["Sex"], person["Province"]+person["City"], person["Signature"], person["RemarkName"]])
	row = 0
	for nickname, sex, city, signature, remarkname in sheetdata:
		cell_format = workbook.add_format()
		cell_format.set_pattern(1)  # This is optional when using a solid fill.
		if row%2 == 0:
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


def exportECharts(_locationStatics, _filename = "echarts.htm"):
	chartsList = []
	for location in _locationStatics:
		chartsList.append({'value': location[1], 'name': location[0]})
	with open(_filename, "w",encoding='utf-8') as f:
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

def exportGroup():
	listg = []
	with open('data.csv','r', encoding="gb18030") as f:
		for line in f:
			row = []
			if line.find("@@") == -1:
				continue
			line = line.split(",")
			for element in line:
				row.append(element)
			listg.append(row)
	if Debug:
		print(list7)
	with open("group.csv","w",encoding="gb18030",newline="") as groupcsv:
		for i in listg:
			response7=i
			response7=str(response7)
			if Debug:
				print(response7)
			groupcsv.write(response7+'\n')
	os.system('call %s' % "group.csv")


def main():
	global Debug
	Debug = False
	cj = http.cookiejar.CookieJar()
	opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
	opener.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0'),
		('Host', 'wx.qq.com'),
		('Accept', 'application/json, text/plain, */*'),
		('Accept-Language', 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2'),
		('Referer', 'https://wx.qq.com/'),
		('DNT', '1')]
	urllib.request.install_opener(opener)
	
	print("获取登录令牌...")
	wxToken["loginQRToken"] = wxGetLoginToken()
	print("获取到登录令牌为:" + str(wxToken["loginQRToken"]))
	
	QRCodeFilename = "qrcode.jpg"
	with open(QRCodeFilename, "wb") as QRCode:
		QRCode.write(wxGetQRCode(wxToken["loginQRToken"]))
	print("已生成二维码图片" + str(QRCodeFilename) + "，请使用手机扫描")
	time.sleep(1)
	os.system('call %s' % QRCodeFilename)

	while True:
		wxLoginStatus = wxGetLoginStatus(wxToken["loginQRToken"])
		if wxLoginStatus[1] == "201":
			print("已经扫描二维码，请在手机上点击登录")
		elif wxLoginStatus[1] == "200":
			print("确认登录微信")
			wxToken["redirectURL"] = getWxRedirectURL(wxLoginStatus[0])
			print("获得返回地址：" + wxToken["redirectURL"])
			break
		time.sleep(1)
	
	wxWebWechatToken = wxRedirect(wxToken["redirectURL"])
	wxToken["skey"] = wxWebWechatToken.split("<skey>")[1].split("</skey>")[0]
	wxToken["wxsid"] = wxWebWechatToken.split("<wxsid>")[1].split("</wxsid>")[0]
	wxToken["wxuin"] = wxWebWechatToken.split("<wxuin>")[1].split("</wxuin>")[0]
	wxToken["pass_ticket"] = wxWebWechatToken.split("<pass_ticket>")[1].split("</pass_ticket>")[0]
	print("获得微信网页版凭据信息：", end="")
	print("我们已经获取您的所有信息")
	wxToken.pop("redirectURL")
	
	wxInitData = json.loads(wxInit(wxToken))
	wxToken["displayname"] = wxInitData["User"]["NickName"]
	wxToken["username"] = wxInitData["User"]["UserName"]
	print(wxToken)
	
	print("\n\n\n\n==========你好，" + wxToken["displayname"] + "！==========")
	print("最近联系人为:")
	for recentCommunicatePerson in wxInitData["ContactList"]:
		displayName = recentCommunicatePerson["RemarkName"] or recentCommunicatePerson["NickName"]
		print("\t" + displayName)
	
	wxContacts = json.loads(wxGetContact(wxToken))
	print("\t目前共有" + str(wxContacts["MemberCount"]) + "位好友")
	
	task = 2133
		
	while task:
		if task == '1':
			print('发送消息')
				#sendAMsg()
			while True:
				contactTo = input("请输入要发送消息的联系人:")
				searchedList = []
				for contacts in wxContacts["MemberList"]:
					if contacts["NickName"].find(contactTo) != -1 or contacts["RemarkName"].find(contactTo) != -1:
						searchedList.append({'NickName': contacts["NickName"], 'RemarkName': contacts["RemarkName"], 'UserName': contacts["UserName"]})
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
			wxSendMsg(wxToken, contactToId, wxMessage)

			
		if task == '2':
			print("正在尝试导出联系人列表")
			exportContect(wxContacts)
			print("已导出联系人列表到friendslist.xlsx")

		if task == '3':
			print("正在针对联系人地域信息进行统计：\n", end="")
			locationStatic = {}
			for person in wxContacts["MemberList"]:
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
			print(locationStatic)
			exportECharts(locationStatic)
			#print(countGroup())

		if task == '4':
			print('导出群组信息')
			exportGroup()
			task = False
		if task == 'tune':
			print('调试')
			Debug = True
			task = 2133
		else:
			print("\n----------1 : 发送一条消息\n----------2 ：导出联系人列表\n----------3 ：统计联系人信息(饼图)\n----------4 ：输出所有群组\n----------tune : 调试模式\n---------- 回车：退出\n")
			task = input("请输入操作编号")



if __name__ == "__main__":
	main()
