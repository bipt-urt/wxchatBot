import WxInterface

wxToken = {}

def main():
	ui = WxInterface.WxInterface()
	
	ui.login()
	
	ui.printRecentContacts()
	
	task = 2133
	while task:
		if task == '1':
			ui.sendMessageUI()
		elif task == '2':
			ui.exportContactUI()
		if task == '3':
			ui.locationStaticUI()
		if task == '4':
			ui.exportGroupUI()
			task = True
		else:
			print("\n----------1 : 发送一条消息\n----------2 ：导出联系人列表\n----------3 ：统计联系人信息(饼图)\n----------4 ：输出所有群组\n----------tune : 调试模式\n---------- 回车：退出\n")
			task = input("请输入操作编号")

if __name__ == "__main__":
	main()