# SingAPu-CheckWordFile
A Visual Basic script for MS Word that checks a word file (in our case an article) for different requirements.

## What it checks
- Check if file extension is ".docx" – if not, output error.
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
- Check if odd numbered instances of **SuS_Kastenheadline** are empty. If so, output an error.
- Check whether paragraphs with format **SuS_Bilddateiname** are always preceded by either a paragraph with the format **SuS_Mengentext** or **SuS_Kastentext** – if not output error.
- Check if paragraph BEFORE + AFTER a paragraph with format **SuS_Kastenheadline** are NOT using **SuS_Kastenheadline** as well – if not output error.
- Check if odd-numbered instance of paragraph with format **SuS_Kastenheadline** include the string "Listing". If it does, check every following paragraph until the next instance of **SuS_Kastenheadline**. If a paragraph is NOT of style **SuS_Quellcode**, output an error in the log.
- Check if paragraphs with format **SuS_Bild/Tabellenunterschrift** are always followed by one of those formats – if not output error:
  - SuS_Mengentext
  - SuS_Kastentext
  - SuS_Absatzheadline
  - SuS_Unter_Absatzheadline
  - SuS_Kasten_Absatzheadline
- Check if content of paragraphs with format **SuS_Bilddateiname** contains a file extension (.tif, .jpg, etc). If NOT, output error. If it does, check if it contains special characters (see filename validation) – if it does, output error.
- Find each odd occurrence of paragraphs with format **SuS_Kastenheadline** and check if they´re are always preceded by a paragraph with the format **SuS_Mengentext** – if not output error.
- Check if paragraph with format **SuS_Bilddateiname** exists. If so, check if filename (without extension) corresponds with image file in Word-doc´s folder. Basically checks if the image files that the .docx file referes to actually exist.
- Check if document only uses styles (character + paragraph) whose name start with "SuS_". If not, output an error.
- Check if paragraphs with the following style contain words/strings that are **italic**:
  - SuS_Headline
  - SuS_Subhead1
  - SuS_Bild/Tabellenunterschrift
  - SuS_Absatzheadlines
  - SuS_Kastenheadline
  - SuS_Tabellenkopf
  - SuS_Quellcode
  - SuS_Quellcode_Kommentar
  - SuS_Links_und_Literatur_Headline
  - SuS_Links_und_Literatur_Text

The script collects all errors in a log file. 

## How to install
### Activate Developer options in MS Word
This only needs to be done once.
1. Open Microsoft Word.
2. Go to FILE > OPTIONS > CUSTOMIZE RIBBON.
3. Under CUSTOMIZE THE RIBBON and under MAIN TABS, select the DEVELOPER check box > press OK.
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image02.png)
### Install / update the script 
1. Go to the github repository.
2. Click on the .vb file.
3. Click button COPY RAW (see screenshot).
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image01.png)
4. Go to the new DEVELOPER tab and hit the VISUAL BASIC button (far left).
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image03.png)
5. In the PROJECT panel (top left-hand side) expand the entry NORMAL > MICROSOFT WORD OBJECTS > double-click THISDOCUMENT.
6. Paste the code from step #3.
7. Click FILE > SAVE NORMAL.
8. Close the Developer view.
### Create shortcut for script
This only needs to be done once.
1. Open Microsoft Word.
2. Go to FILE > OPTIONS > CUSTOMIZE RIBBON.
3. Switch to QUICK ACCESS TOOLBAR (see screenshot – step 1).
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image04.png)
4. In the dropdown menu choose MAKROS (see screenshot – step 2).
5. Select the script based on it´s name (see screenshot – step 3).
6. Click the right-arrow button to add it to the quick access toolbar (see screenshot – step 4).

The shortcut has been created and can be used (see screenshot).
![alt text](https://github.com/mschuetze/SingAPu-CheckWordFile/blob/main/image05.png)