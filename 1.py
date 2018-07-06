import WxBot
import time
import json
import os

wxToken = {}

def main():
	global Debug
	Debug = False
	bot = WxBot.WxBot()
	
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
