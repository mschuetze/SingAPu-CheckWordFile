# SingAPu-CheckWordFile
A Visual Basic script for MS Word that checks a word file (in our case an article) for different requirements.

## What it checks
- Check whether the first paragraph contains a pipe (‘’|”"). If not, an error is output. If it does, the system checks whether the first paragraph has the paragraph format **SuS_Mengentext** - if not, convert automatically.
- Check whether second paragraph has paragraph format **SuS_Headline** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Headline** - if not, output error.
- Check whether third paragraph has format **SuS_Subhead1** - if not, convert automatically.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Subhead1** - if not, output error.
- Check whether 0 OR 1 paragraph has the paragraph format **SuS_Subhead2** - if not, output error.
- Check whether exactly 1 paragraph has the paragraph format **SuS_Autorname** - if not, output error.
