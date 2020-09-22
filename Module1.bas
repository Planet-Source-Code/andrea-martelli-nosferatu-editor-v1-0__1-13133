Attribute VB_Name = "Module1"
'******************************'
'*     Nosferatu Editor v1.0  *'
'******************************'
'*    author:Andrea Martelli  *'
'******************************'
'*                            *'
'* you can download this code *'
'* and many others visiting   *'
'* Visual Basic Italia        *'
'******************************'
'http://digilander.iol.it/VBItalia/INDICE.html

Option Explicit
'comincia l'elenco delle costanti
Public Enum RasterOps
'copia l'immagine primitiva nell'immagine di destinazione
SRCCOPY = &HCC0020
'combina i pixel dell'immagine di destinazione con quelli 'dell'immagine di partenza usando l'operatore Booleano AND
SRCAND = &H8800C6
'combina i pixel dell'immagine di destinazione con quelli 'dell'immagine di partenza usando l'operatore Booleano XOR
SRCINVERT = &H660046
'combina i pixel dell'immagine di destinazione con quelli 'dell'immagine di partenza usando l'operatore Booleano OR
SRCPAINT = &HEE0086
'inverte l'immagine di destinazione e la combina con l'immagine di partenza usando l'operatore Booleano AND
SRCERASE = &H4400328
'operazioni sull'output
WHITENESS = &HFF0062
BLACKNESS = &H42
End Enum
'dichiarazioni pubbliche della funzione BitBlt
Declare Function BitBlt Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal dwRop As RasterOps _
) As Long


