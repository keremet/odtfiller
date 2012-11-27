..\odtfiller.exe xmlfile.xml template.odt . out.odt
del content.xml
unzip out.odt content.xml
diff content.xml etalon_content.xml
