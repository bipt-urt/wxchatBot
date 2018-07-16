

import hashlib
with open('pic.jpg','rb') as f:
	hashm = hashlib.new('md5', f.read()).hexdigest()
	print(hashm)
