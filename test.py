

def flt_from_str(string):
    """
    :param string:строка должна передаваться в формате '51.660781, 39.200296'
    :return: tuple (x_coord,y_coord)
    """
    str_x_coord=''
    str_y_coord=''
    for i in range(0,len(string)):
        if string[i]!=',':
            str_x_coord+=string[i]
        else:
            str_y_coord=string[i+2:]
            coordinates=(float(str_x_coord),float(str_y_coord))
            return coordinates

# k=flt_from_str('53.941622, 78.467615')
# print(type(k))

dict