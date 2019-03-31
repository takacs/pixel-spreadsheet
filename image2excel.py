import xlsxwriter
from PIL import Image
import sys

def toHexa(val, channel):
    """
    Convert decimal value to hexadecimal color code where
    only channel given has a value and the other two are set
    to 00.

    val: decimal value
    channel: color channel
    """

    if channel == 'r':  
        return '#' + '{:02x}'.format(val) + '0000'
    elif channel == 'b':
        return '#00' + '{:02x}'.format(val) + '00'
    else:
        return '#0000' + '{:02x}'.format(val)


def main():

    # load image
    im = Image.open(sys.argv[1])
    im.thumbnail((128,128), Image.ANTIALIAS) # downsampling, aspect ratio stays the same
    pix = im.load()
    imsize = im.size
    
    # create excel workbook/worksheet
    workbook = xlsxwriter.Workbook( './output.xls')
    worksheet = workbook.add_worksheet('image')

    # color excel cells
    for x in range(imsize[0]):
        for y in range(0,imsize[1]*3, 3):
            colors = pix[x, y/3]
            for i,channel in enumerate(['r','g','b']):
                hexcode = toHexa(colors[i], channel=channel)
                wbformat = workbook.add_format({'bg_color':hexcode})
                worksheet.write(y+i,x,hexcode, wbformat)

    workbook.close()

if __name__ == '__main__':
    main()
