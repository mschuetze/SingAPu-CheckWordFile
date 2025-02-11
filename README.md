# SingAPu-CheckWordFile
A Visual Basic script for MS Word that checks a word file (in our case an article) for different requirements.

## What it checks
- Validates the .docx filename. If special characters are included, the script stops and throws an error. Special characters are:
    - German Umlaute
    - space
    - !
    - @
    - \#
    - $
    - %
    - ^
    - &
    - \*
    - ( and )
    - \+
    - =
    - { and }
    - [ and ]
    - |
    - \ and /
    - :
    - ;
    - ""
    - '
    - < and >
    - ,
    - .
    - ?
    - ~
    - `
- Check whether the first paragraph contains a pipe (‘’|”"). If not, an error is output. If it does, the system checks whether the first paragraph has the paragraph format **SuS_Mengentext** - if not, convert automatically.
- Check whether second paragraph has paragraph format **SuS_Headline** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Headline** - if not, output error.
- Check whether third paragraph has format **SuS_Subhead1** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Subhead1** - if not, output error.
- Check whether 0 OR 1 paragraph has the paragraph format **SuS_Subhead2** - if not, output error.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Autorname** - if not, output error.
- Check if paragraphs with format **SuS_Subhead2** are always followed by a paragraph with format **SuS_Autorname**.
- Check whether 0 OR 1 paragraph has the paragraph format **SuS_Links_und_Literatur_Headline** - if not, output error.
- Check if paragraph with format **SuS_Links_und_Literatur_Headline** is always followed by a paragraph with format **SuS_Links_und_Literatur_Text**.
- Check if paragraphs with format **SuS_Links_und_Literatur_Text** are always followed by a paragraph with format **SuS_Links_und_Literatur_Text**.
- Check if paragraphs with format **SuS_Bilddateiname** are always followed by either a paragraph with format **SuS_Bild/Tabellenunterschrift** or **SuS_Autor_Kurzbiografie**.
- Check if the number of paragraphs with paragraph format **SuS_Kastenheadline** is an integer multiple of 2 (each box is opened + closed). If this is not the case, output an error.
- Check whether paragraphs with format **SuS_Bilddateiname** are always preceded by either a paragraph with the format **SuS_Mengentext** or **SuS_Kastentext** – if not output error.
- Check if paragraph BEFORE + AFTER a paragraph with format **SuS_Kastenheadline** are NOT using **SuS_Kastenheadline** as well – if not output error.
- Check if paragraphs with format **SuS_Bild/Tabellenunterschrift** are always followed by one of those formats – if not output error:
  - SuS_Mengentext
  - SuS_Kastentext
  - SuS_Absatzheadline
  - SuS_Unter_Absatzheadline
- Check if content of paragraphs with format **SuS_Bilddateiname** contains special characters (see filename validation) – if it does, output error.

The script collects all errors in a log file. 

## How to install
1. Go to the github repository.
2. Click on the .vb file
3. Click button COPY RAW (see screenshot)
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image01.png)
4. Open Microsoft Word.
5. Go to FILE > OPTIONS > CUSTOMIZE RIBBON.
6. Under CUSTOMIZE THE RIBBON and under MAIN TABS, select the DEVELOPER check box > press OK.
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image02.png)
7. Go to the new DEVELOPER tab and hit the VISUAL BASIC button (far left).
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image03.png)
8. In the PROJECT panel (top left-hand side) expand the entry NORMAL > MICROSOFT WORD OBJECTS > double-click THISDOCUMENT.
9. Paste the code from step #3.
10. Click FILE > SAVE NORMAL  
