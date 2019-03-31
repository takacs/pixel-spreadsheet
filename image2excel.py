#!/usr/bin/env python3
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

    if len(sys.argv) < 2:
        sys.exit('No input image!')
    # load image
    im = Image.open(sys.argv[1])
    imsize = im.size
    print(f'Input image size: ({imsize[0]}, {imsize[1]})')
    if imsize[0] > 128 or imsize[1] > 128:
        print('Downsampling image.')
        im.thumbnail((128,128), Image.ANTIALIAS) # downsampling, aspect ratio stays the same
        imsize = im.size
        print(f'Image size after downsampling: ({imsize[0]}, {imsize[1]})')
    pix = im.load()
    
    # create excel workbook/worksheet
    workbook = xlsxwriter.Workbook( './output.xls')
    worksheet = workbook.add_worksheet('image')

    # color excel cells
    for x in range(imsize[0]):
        for y in range(0, imsize[1]*3, 3):
            colors = pix[x, y/3]
            for i,channel in enumerate(['r','g','b']):
                hexcode = toHexa(colors[i], channel=channel)
                wbformat = workbook.add_format({'bg_color':hexcode})
                worksheet.write(y+i,x,colors[i], wbformat)

    workbook.close()
    print('Spreadsheet complete. Saved to ./output.xls')

if __name__ == '__main__':
    main()
