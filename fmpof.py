
from __future__ import print_function
import builtins
import past
import six
import sys
import os
import getpass
import argparse
import struct
import base64
import zipfile
from Crypto.Cipher import AES
from Crypto import Random

# Constants
salt_size = 12
io_code = 'utf-8'     # FIXME: May not be good for Windows, but I use only ASCII so I leave it for now.


def error_exit(message):
    print("ERROR: " + message, file=sys.stderr)
    sys.exit(2)


def warning(message):
    print("WARNING: " + message, file=sys.stderr)


def error_exit_if_not_on_win32():
    if sys.platform != "win32":
        print("This program does not run on Linux or macOS,")
        print("because it depends on pywin32 module that runs only on Windows.")
        sys.exit(2)


def error_exit_if_pywin32_is_missing():
    try:
        import win32com.client
    except:
        print("It seems win32com module is missing.")
        print("Install pywin32 module, and try again.")
        print("eg) pip install pywin32")
        sys.exit(2)


def encode_and_pad_for_aes_key(raw_password, salt=None):
    """
        AES requires a key of 16, 24, or 32 bytes long.
        Also, raw_password is a string but the AES key must be bytes.
        It does conversion here.
    """
    binary_password = raw_password.encode(io_code)
    if salt != None:
        binary_password = salt + binary_password
    l = len(binary_password)
    if l > 32:
        warning("Password is truncated because it is too long.")
        return binary_password[0:32]
    if l > 24:
        return binary_password + b'\0' * (32 - l)
    if l > 16:
        return binary_password + b'\0' * (24 - l)
    return binary_password + b'\0' * (16 - l)


def init_db(file_name):
    if os.path.exists(file_name):
        error_exit("'%s' already exists" % file_name)
    try:
        with open(file_name, "w") as f:
            pass
    except:
        error_exit("Cannot open '%s' for write" % file_name)


def get_passwords(file_name, password):
    retvals = []
    try:
        with open(file_name, "r") as f:
            for l in f:
                decoded_bin = base64.b64decode(l.encode('ascii'))
                salt, encrypted_pass = struct.unpack("<%ds256s" % salt_size, decoded_bin)
                aesobj = AES.new(encode_and_pad_for_aes_key(password, salt), AES.MODE_ECB)
                decrypted_packed_pass = aesobj.decrypt(encrypted_pass)
                decrypted_pass, = struct.unpack("256p", decrypted_packed_pass)
                try:
                    retvals.append(decrypted_pass.decode(io_code))
                except UnicodeDecodeError:
                    error_exit("Unicode error. Probably you input a wrong password.")
    except IOError:
        error_exit("Cannot open '%s'" % file_name)
    return retvals


def list_passwords(file_name, password):
    passwords = get_passwords(file_name, password)
    for p in passwords:
        print(p)


def add_password(file_name, password, new_pass_line):
    try:
        with open(file_name, "a") as f:
            if new_pass_line == None:
                new_pass_line = getpass.getpass("New password (of MS Office files): ")
                new_pass_line2 = getpass.getpass("Reenter: ")
                if new_pass_line != new_pass_line2:
                    error_exit("Does not match")
            binary_new_pass = new_pass_line.encode(io_code)
            if 250 < len(binary_new_pass):
                error_exit("New password (of MS Office files) is too long (>250 bytes)")
            rndfile = Random.new()
            salt = rndfile.read(salt_size)
            aes_key = encode_and_pad_for_aes_key(password, salt)
            # print("aes_key length = %d" % len(aes_key))
            aesobj = AES.new(aes_key, AES.MODE_ECB)
            packed_pass = struct.pack("256p", binary_new_pass)
            # print("packed_pass length = %d" % len(packed_pass))
            encrypted_key = aesobj.encrypt(packed_pass)
            output_binary = salt + encrypted_key
            # print("output binary length = %d" % len(output_binary))
            print(base64.b64encode(output_binary).decode('ascii'), file=f)
    except IOError:
        error_exit("Cannot open '%s'" % file_name)


def get_file_type(file_name):
    if file_name.endswith(".xls"): return "excel"
    if file_name.endswith(".xlsx"): return "excel"
    if file_name.endswith(".doc"): return "word"
    if file_name.endswith(".docx"): return "word"
    if file_name.endswith(".ppt"): return "ppt"
    if file_name.endswith(".pptx"): return "ppt"
    if file_name.endswith(".zip"): return "zip"
    error_exit("Unknown file type for file '%s'" % file_name)


def open_file(msoffice_file_name, db_file_name, raw_password, args):
    passwords = get_passwords(db_file_name, raw_password)
    file_type = get_file_type(msoffice_file_name)

    import win32com.client
    if file_type == "excel":
        office_obj = win32com.client.gencache.EnsureDispatch('Excel.Application')
    elif file_type == "word":
        office_obj = win32com.client.gencache.EnsureDispatch('Word.Application')
    elif file_type == "ppt":
        office_obj = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
    elif file_type == "zip":
        office_obj = None
    else:
        error_exit("Logic error")

    worked = False
    num_tested = 0
    for p in passwords:
        num_tested += 1
        print("%d passwords tested" % num_tested, end="\r")
        # if True:
        try:
            if file_type == "excel":
                workbook = office_obj.Workbooks.open(msoffice_file_name, False, False, None, p)
            elif file_type == "word":
                document = office_obj.Documents.Open(msoffice_file_name, False, False, False, p)
            elif file_type == "ppt":
                presentaion = office_obj.Presentations.open(msoffice_file_name + ":" + p + "::")
            elif file_type == "zip":
                zipfile = zipfile.ZipFile.open(msoffice_file_name, mode, p)
            else:
                error_exit("Logic error")
            print("")
            if args.show:
                print("'%s' was the password." % p)
            else:
                print("Tada! It opened.")
            if file_type != "zip":
                office_obj.Visible = True
            worked = True
            break
        except win32com.client.pywintypes.com_error:
            # print("%s did not work" % p)
            pass
    if not worked:
        print("")
        error_exit("All passwords tested did not match.")


def main():
    error_exit_if_not_on_win32()
    error_exit_if_pywin32_is_missing()

    parser = argparse.ArgumentParser(description='Opens an MS Office file with automatic decryption.')
    parser.add_argument('file', type=str, nargs='?', help='MS Office file (doc/docx/xls/xlsx/ppt/pptx)')
    parser.add_argument('--password', type=str, help='DB password. If not given, we ask you')
    parser.add_argument('--db', type=str, help='Password database file (default: ~/.msopass)')
    parser.add_argument('--list', action='store_true', help='List passwords')
    parser.add_argument('--add', action='store_true', help='Add a new password')
    parser.add_argument('--addpass', help='A new password (only valid with --add)')
    parser.add_argument('--init', action='store_true', help='Initialize the db file')
    parser.add_argument('--show', action='store_true', help='Show the correct password if opened')

    args = parser.parse_args()

    if args.db:
        db_file = args.db
    else:
        db_file = os.environ["HOMEDRIVE"] + os.environ["HOMEPATH"] + "\\.msopass"
    # print("DB file = %s" % db_file); sys.exit(0)

    if args.init:
        init_db(db_file)
        sys.exit(0)

    if args.password:
        raw_password = args.password
    else:
        raw_password = getpass.getpass("DB password: ")

    if args.list:
        list_passwords(db_file, raw_password)
        sys.exit(0)

    if args.add:
        add_password(db_file, raw_password, args.addpass)
        sys.exit(0)

    # print(args.file); sys.exit(0)
    open_file(args.file, db_file, raw_password, args)


if __name__ == "__main__":
    main()
