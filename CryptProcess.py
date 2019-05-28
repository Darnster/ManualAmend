"""
Module to encrypt/decrypt passwords to store in
"""

from cryptography.fernet import Fernet
import sys

def encrypt(message: bytes, key: bytes) -> bytes:
    return Fernet(key).encrypt(message)

def decrypt(token: bytes, key: bytes) -> bytes:
    return Fernet(key).decrypt(token)

def generateKey():
    key = Fernet.generate_key()
    return key

if __name__ == "__main__":
    pwd = sys.argv[1]
    print("Password = %s" % pwd)
    key = generateKey()
    print("your key:\n%s" % key.decode() )
    print("(Keep this in a safe place and pass it in to the Manual Amends script)\n")
    enCryptedPWD = encrypt( pwd.encode(), key )
    print("your encrypted password:\n%s" % enCryptedPWD.decode() )
    print("(Store this in your config file as the pwd)\n" )
    plain = decrypt( enCryptedPWD, key).decode()
    print("Roundtrip testing - your decrypted password:\n%s" % plain)
    print("Printed purely to re-assure you that it will decrypt as expected :-)")


"""
>>> key = Fernet.generate_key()
>>> print(key.decode())
itYUevR5OTOHVQG8KcI1_fPwnoYahH33X-RCxRa6mwU=
>>> message = 'John Doe'
>>> encrypt(message.encode(), key)
'gAAAAABciT3pFbbSihD_HZBZ8kqfAj94UhknamBuirZWKivWOukgKQ03qE2mcuvpuwCSuZ-X_Xkud0uWQLZ5e-aOwLC0Ccnepg=='
>>> token = _
>>> decrypt(token, key).decode()
'John Doe'
"""
