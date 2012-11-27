..\..\odtfiller.exe xmlfile.xml template.ods . out.ods
del content.xml
unzip out.ods content.xml
diff content.xml etalon_content.xml
