from excelutils import read_dict, write
import re

if __name__ == "__main__":
    data = read_dict("ECN.xls")
    data_msg = [['ASIN','item_name', 'Mark', 'Keywords', 'Attributes', 'Attribute Values']]
    for i in data:
        ASIN= i['ASIN']
        item_name = i['item_name']
        Mark = i['Mark']
        Attributes = i['Attributes']
        name = i['Keywords']
        value = i['Attribute Values']
        msg = re.findall(r'[^。]*?{}[^。]*?。'.format(name), value)
        data_msg.append([ASIN, item_name, Mark, name, Attributes, msg])
    write(path="target.xls", rows=data_msg)
