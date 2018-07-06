class WxBot:
	def __init__(self):
		pass
	
	def getR(self):
		import random
		randomTicket = str(random.random())[2:] + '1'
		return randomTicket
	
	