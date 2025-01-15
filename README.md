# SingAPu-CheckWordFile
A Visual Basic script for MS Word that checks a word file (in our case an article) for different requirements.

## What it checks
- Validates the .docx filename. If special characters are included, the script stops and throws an error.
    - German Umlaute not yet included
- Check whether the first paragraph contains a pipe (‘’|”"). If not, an error is output. If it does, the system checks whether the first paragraph has the paragraph format **SuS_Mengentext** - if not, convert automatically.
- Check whether second paragraph has paragraph format **SuS_Headline** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Headline** - if not, output error.
- Check whether third paragraph has format **SuS_Subhead1** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Subhead1** - if not, output error.
- Check whether 0 OR 1 paragraph has the paragraph format **SuS_Subhead2** - if not, output error.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Autorname** - if not, output error.
- Check if paragraphs with format **SuS_Subhead2** are always followed by a paragraph with format **SuS_Autorname**.
- Check if paragraphs with format **SuS_Bilddateiname** are always followed by either a paragraph with format **SuS_Bild/Tabellenunterschrift** or **SuS_Autor_Kurzbiografie**.
- Check if the number of paragraphs with paragraph format **SuS_Kastenheadline** is an integer multiple of 2 (each box is opened + closed). If this is not the case, output an error.
- Check whether paragraphs with format **SuS_Bilddateiname** are always preceded by either a paragraph with the format **SuS_Mengentext** or **SuS_Kastentext**.

The script collects all errors in a log file. 
