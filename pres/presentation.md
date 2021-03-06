### Jegyzet

## A problema

[https://index.hu/mindekozben/poszt/2019/03/26/ha_mindig_is_fotot_akart_excel-tablazatta_konvertalni_ne_keressen_tovabb](https://index.hu/mindekozben/poszt/2019/03/26/ha_mindig_is_fotot_akart_excel-tablazatta_konvertalni_ne_keressen_tovabb)

[http://www.think-maths.co.uk/spreadsheet](http://www.think-maths.co.uk/spreadsheet)


* Az elektronikus kijelzok valojaban tablazatok, ahol a pixelek mindossza a piros, a zold es a kek szinek eltero fenyereju keppontjait jelenitik meg.

<img src="actualpixels.png" width=100 height=100>

* Otlet: Excel-ben is lehetne abrazolni ezeket a keppontokat ugy, hogy az egyes cellak hatterenek a piros, a zold vagy a kek csatornak megfelelo arnyalatait rendeli.

* Egy pixel:

<img src="apixelinexcel.png" width=100>


* Az eredmeny: 

<img src="indexcelmage.png" width=450>

## A feladat

* Atultetni Python-ba.
* Valahogy ki kell olvasni a pixelek ertekeit, valamint valahogy irni kene egy excel tablazatba.


### Pillow konyvtar

[docs](https://pillow.readthedocs.io/en/stable/)

* Kepek megnyitasara, manipulasara, kulonbozo formatumokba mentesere irt konyvtar.

#### Egy kep betoltese

```
im = Image.open(sys.argv[1])
imsize = im.size
pix = im.load()
```

### XlsxWriter konyvtar

[docs](https://xlsxwriter.readthedocs.io/)

* Excel fileok manipulalasara irt konyvtar.

```
workbook = Workbook('./output.xls')
worksheet = workbook.add_worksheet('image')
wbformat = workbook.add_format({'bg_color':colorcode})
worksheet.write(row, column, text, format)
workbook.close()
```


### Nagy kepek lassan toltenek be

* Tul nagy kepek lassan vagy egyaltalan nem toltottek be excelbe, ezert ezeknek a kepeknek a meretet kisebbre vettem.

```
im.thumbnail((128,128))
```


## Pelda

### Bemenet

<img src="lena.png" width=400>

### Kimenet

<img src="lenaexcel.JPG" width=400>