import sys

def open_file(filename):
    list = []
    try:
        f = open(filename, 'rU')
        for str in f:
            list.extend(str.split())
        f.close()
    except FileNotFoundError:
        print("You must create file %s. Then input username in 1st line and password in 2nd line" % filename)
        sys.exit(1)
    return list


def check_send_notice():
    last_value = open_file(".send_email")[0]
    try:
        f = open(".send_email", 'w+')
        if last_value == '0':
            f.write('1')
        else:
            f.write('0')
        f.close()
    except FileNotFoundError:
        print("You must create file %s. Then input username in 1st line and password in 2nd line")
        sys.exit(1)

    if last_value == '1':
        return True
    else:
        return False
