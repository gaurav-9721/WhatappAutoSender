import pandas as pd
import pywhatkit

from set_data import excel_path, caption_template_path, field_caption, tab_close_time, image_wait_time, image_path

exceldf = None


def get_blanks(ind):

    global exceldf
    blanks = []
    for f in field_caption:
        field = exceldf[f][ind]
        blanks.append(field)
    return blanks

def send_msg(index, caption):

    global exceldf
    caption = caption.format(*get_blanks(index))

    contact = '+91' + str(exceldf['Contact'][index]).strip()
    try:

        pywhatkit.sendwhats_image(contact, image_path,caption, image_wait_time, True, 5)
        print('Image with caption sent to', contact)
    except:
        print('Error! Cannot send Image to', contact)


if __name__ == '__main__':
    exceldf = pd.read_excel(excel_path)

    file = open('caption_template.txt')
    message = file.read()
    file.close()

    for ind in exceldf.index:
        send_msg(ind, message)


