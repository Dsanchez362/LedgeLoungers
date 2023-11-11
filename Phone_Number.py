import re

# This function iterates over specific instructions looking for a phone number. Return True when all peramiters are met. Return False when they are not
def isPhoneNumber(text):
    if len(text) != 12:
        return False
    for i in range (0,3):
        if not text[i].isdecimal():
            return False
        if text[3] != '-':
            return False
        for i in range(4,7):
            if not text[i].isdecimal():
                return False
        if text[7] != '-':
            return False
        for i in range(8,12):
            if not text[i].isdecimal():
                return False
        return True

# This additional code looks for a phone number, using the above code and adding the new code
message = 'Call me at 415-555-1011 tomorrow. 415-555-9999 is my office.'
for i in range(len(message)):
    chunk = message[i:i+12]
    if isPhoneNumber(chunk):
        print('Phone Number Found: ' + chunk)
print('Done')

# This code uses Regex to look for phone numbers without having to be specific on the format
phoneNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d\d')
mo = phoneNumRegex.search('My number is 415-555-4242')
print('Phone number found: ' +mo.group())


