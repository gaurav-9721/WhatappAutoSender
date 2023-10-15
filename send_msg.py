import pandas as pd
import pywhatkit
from set_data import excel_path, message_template_path, field_msg, tab_close_time, wait_time

exceldf = None


def get_blanks(ind):

    global exceldf
    blanks = []
    for f in field_msg:
        field = exceldf[f][ind]
        blanks.append(field)
    return blanks

def send_msg(index, message):

    global exceldf
    message = message.format(*get_blanks(index))

    contact = '+91' + str(exceldf['Contact'][index]).strip()
    try:

        pywhatkit.sendwhatmsg_instantly(contact, message, wait_time, True, tab_close_time)
        print('Message sent to', contact)
    except:
        print('Error! Cannot send message to', contact)




if __name__ == '__main__':
    exceldf = pd.read_excel(excel_path)

    file = open(message_template_path)
    message = file.read()
    file.close()

    for ind in exceldf.index:
        send_msg(ind, message)


