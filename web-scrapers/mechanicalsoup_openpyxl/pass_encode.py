import getpass
def simple_encode(vl):
	result = ""
	for ch in vl:
		result += chr(ord(ch) - 15)
	return result
	
password = getpass.getpass("Please enter your password: ")
with open("password", "w") as f:
	f.write(simple_encode(password))