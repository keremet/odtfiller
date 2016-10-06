/*
TODO: новый входной формат, закрытие файлов, регрессионные тесты (Unix и windows)
*/
#include <stdio.h>
#include <stdlib.h>
#include <iconv.h>
#include <fcntl.h>
#include <errno.h>
#include <sys/types.h>
#include <sys/stat.h>
extern "C"{
#include <zip.h>
}
#ifdef _WIN32
	#include <windows.h>
	#include <shellapi.h>
	#define SLASH '\\'
#else	
	#define SLASH '/'
#endif

#include <string>
#include <cstring>
/*#include <sstream>
#include <list>*/
using namespace std;

void replaceAll(string *content, const string key, const string value){
	for(string::size_type pos=0;(pos = content->find(key, pos))!=string::npos;){
		content->replace(pos,key.length(), value);
	}
}

void replaceField(string *content, const string key, const string value){
	fprintf(stderr,"|%s|%s|\n", key.c_str(), value.c_str());
	/*string field = "<text:user-field-get text:name=\""+key+"\">"+key+"</text:user-field-get>";	
	replaceAll(content, field, value); */
	string fldBegin = "<text:user-field-get text:name=\""+key+"\">";
	for(string::size_type pos=0;(pos = content->find(fldBegin, pos))!=string::npos;){
		string::size_type fldEndPos = content->find("</text:user-field-get>", pos);
		content->replace(pos,fldEndPos+22 /*strlen("</text:user-field-get>")*/ - pos, value);
	}		
}

void replaceFieldWithParagraph(string *content, const string key, const string value){
	//fprintf(stderr,"|%s|%s|\n", key.c_str(), value.c_str());
	string field = "<text:user-field-get text:name=\""+key+"\">"+key+"</text:user-field-get>";	
	for(string::size_type pos=0;(pos = content->find(field, pos))!=string::npos;){
		string::size_type paraEndPos = content->find("</text:p>", pos)+9; //strlen("</text:p>");
		string::size_type paraBeginPos = content->rfind("<text:p ", pos); //strlen("</text:p>");
		content->replace(paraBeginPos,paraEndPos-paraBeginPos, value);
	}
}

void addStyles(string *content, const string value){
	string::size_type stylesPos = content->find("</office:automatic-styles>");
	if(stylesPos!=string::npos)
		content->insert(stylesPos,value);
}

void addOfficeText(string *content, const string value){
	content->insert(content->find("</office:text>"),value);
}
/*
void refreshTextFields(string *content, const string value){
	list<string> textFieldsList;
	for(string::size_type pos=0;(pos = content->find("<text:bookmark-start text:name=\"", pos))!=string::npos;){
		pos+=32; //strlen("<text:bookmark-start text:name=\"")
		string name= content->substr(pos , content->find("\"/>", pos)-pos);
		textFieldsList.push_back (name);
	}
	for (list<string>::iterator it = textFieldsList.begin(); it != textFieldsList.end(); it++){
		fprintf(stderr,"FieldNameL=%s\n", it->c_str());
		string::size_type bookmarkPos = content->find("<text:bookmark-start text:name=\"" + *it +"\"/>"); 
		if(bookmarkPos==string::npos){
			fprintf(stderr,"BookMark=%s not found\n", it->c_str());
			continue;
		}		
		string::size_type listRootPos = content->rfind("<text:list xml:id=\"", bookmarkPos);
		if(listRootPos==string::npos){
			fprintf(stderr,"List root for BookMark=%s not found\n", it->c_str());
			continue;
		}
		ostringstream oss;
		for(string::size_type pos=listRootPos;(pos = content->find("<text:list>", pos+1))!=string::npos;){
			oss<<'.';
		}
		fprintf(stderr,"value=%s\n", oss.str().c_str());
		//replaceAll(content, const string key, oss.str())
	}
}
*/

void addRow(string *content, const string key, const string value){
	//printf("%s\n", key.c_str());
	string::size_type tablePos = content->find("<table:table table:name=\"" +key+ "\"");
	if(tablePos!=string::npos){
//		printf("%i\n", tablePos);
		string::size_type tableEndPos = content->find("</table:table>", tablePos);
//		printf("%i\n", tableEndPos);
		if(tableEndPos!=string::npos)
			content->insert(tableEndPos,value);
	}
}

void insertRow(string *content, const string key, string value){
	//printf("%s\n", key.c_str());
	string::size_type tablePos = content->find("<table:table table:name=\"" +key+ "\"");
	if(tablePos!=string::npos){
		string::size_type tableEndPos = content->find("</table:table>", tablePos);
		if(tableEndPos!=string::npos){
			int rowsBefore=atoi(value.c_str());
			value=value.substr(value.find(' ')+1);
			string::size_type rowPos = content->find("<table:table-row", tablePos);
			if(rowPos==string::npos) return;
			for(;rowsBefore;rowsBefore--){
				rowPos = content->find("</table:table-row>", rowPos+1);
				if(rowPos==string::npos) return;
				rowPos += 18 /*strlen("</table:table-row>")*/ ;
			}
			//printf("%i\n\n%s", rowsBefore, value.c_str());
			if(rowPos>tableEndPos) return;
			content->insert(rowPos,value);
		}
	}
}


void renameStylesAndObjects(string *content, const string prefix){	
	for(string::size_type stylePos = 0;(stylePos = content->find("<style:style style:name=\"", stylePos))!=string::npos;stylePos+=26){
		string::size_type styleNamePos=stylePos+25;//strlen("<style:style style:name=\"");
		string styleName=content->substr(styleNamePos, content->find('"',styleNamePos)-styleNamePos);
		replaceAll(content, string("name=\"")+styleName+'"', string("name=\"")+prefix+'_'+styleName+'"');
	}
}

void deleteTable(string *content, const string tableName){	
	string::size_type tablePos = content->find("<table:table table:name=\"" +tableName+ "\"");
	if(tablePos!=string::npos){
		string::size_type tableEndPos = content->find("</table:table>", tablePos);
		if(tableEndPos!=string::npos){
			content->erase(tablePos, tableEndPos+14 /*strlen("</table:table>")*/ - tablePos);
		}
	}
}


string path2copyfile_templates;

string getContent(const char* odtFileName){
	struct zip *zs;
	int i, idx, err;
	char errstr[1024];
    if ((zs=zip_open((path2copyfile_templates+odtFileName).c_str(), 0, &err)) == NULL) {
		zip_error_to_str(errstr, sizeof(errstr), err, errno);
		fprintf(stderr, "cannot open  %s: %s\n",odtFileName,errstr);
		return "";
    }	
    if ((idx=zip_name_locate(zs, "content.xml", ZIP_FL_NODIR|ZIP_FL_NOCASE)) != -1) {
		struct zip_file *zf;
	//Чтение содержимого		
		string content="";
		if ((zf=zip_fopen_index(zs, idx, 0)) == NULL) {
			fprintf(stderr, "cannot open file %i in archive: %s\n",idx,zip_strerror(zs));
			return "";
		}
		int n;
		char buf[8192];
		while ((n=zip_fread(zf, buf, sizeof(buf))) > 0) {
			//~ for(int i=0;i<n;i++)
				//~ cout << buf[i];
			for(int i=0;i<n;i++)
				content+=buf[i];
		}		
		if (n < 0) {
			fprintf(stderr, "error reading file %i in archive: %s\n",idx,zip_file_strerror(zf));
			zip_fclose(zf);
			return "";
		}
		zip_fclose(zf);
		return content;
	}else{
		fprintf(stderr, "zip_name_locate error\n");
	}	
	return "";
}

string getTextBetween(string *content, const char* begin, const char* end){
	string::size_type  posB = content->find(begin);
	if(posB==string::npos) return "";
	posB = content->find('>',posB);
	if(posB==string::npos) return "";
	posB++;
	string::size_type  posE=content->find(end,posB); 
	if(posE==string::npos) return "";
	return content->substr(posB,posE-posB); //Стили
}

string getOfficeText(string *content){	
	return getTextBetween(content, "<office:text", "</office:text>");
}

string getStylesText(string *content){
	return getTextBetween(content, "<office:automatic-styles", "</office:automatic-styles>");
}

void parseAndReplace(string *content, const string xmlStr){
	string office_text = getOfficeText(content);
	//replaceAll(&xmlStr, "<remove_me/>\x0A", ""); - comments in xmlStr
	//if(office_text.empty()) return;
	
	for(string::size_type pos=0;(pos = xmlStr.find("<Field Name=\"", pos))!=string::npos;){
		string key, value;
		string::size_type pos0=pos+13; //strlen("<Field Name=\"");
		if((pos = xmlStr.find("\">", pos0))!=string::npos){
			key=xmlStr.substr(pos0, pos-pos0);
		}else break;
		pos0=pos+2; //strlen("\">")
//		fprintf(stderr,"xmlStr[pos0]==%i\nxmlStr[pos0+1]==%i\n", xmlStr[pos0], xmlStr[pos0+1]);
		while (xmlStr[pos0]==0xD) pos0+=2; //Есть возможность попасть на перенос строки - надо ее учесть
		if((pos = xmlStr.find("</Field>", pos0))!=string::npos){
			value=xmlStr.substr(pos0, pos-pos0);
		}else break;
		string::size_type posTO;
		if((posTO=key.find("\" TypeOut=\"code"))!=string::npos){
			bool fAddRow,fInsertRow=false;
			if(!(fAddRow = (key.find("_addrow")!=string::npos))){
				fInsertRow = (key.find("_insertrow")!=string::npos);
			}
			key=key.substr(0, posTO);			
			replaceAll(&value, "&lt;","<");
			replaceAll(&value, "&gt;",">");
			replaceAll(&value, "&quot;","\"");
//			fprintf(stderr,"TypeOut\n");
			if(fInsertRow){
//				fprintf(stderr,"TypeOut1\n");
				insertRow(content, key, value);
			}else if(fAddRow){
//				fprintf(stderr,"TypeOut2\n");
				addRow(content, key, value);
			}else if(key == "openoffice_styles_list"){
//				fprintf(stderr,"TypeOut3\n");
				addStyles(content, value);
			}else
				replaceFieldWithParagraph(content, key, value);
//			fprintf(stderr,"TypeOut4\n");				
			continue;			
    }
		if((posTO=key.find("\" TypeOut=\"copyfile"))!=string::npos){			
			string odtFileContent = getContent(value.c_str()); //Извлечь content.xml из odt			
			renameStylesAndObjects(&odtFileContent, value);		
			//printf("%s\n", odtFileContent.c_str());
			//Выделить текст и стили
			string external_office_text = getOfficeText(&odtFileContent);
			string external_styles_text = getStylesText(&odtFileContent);
			if(external_office_text.empty() || external_styles_text.empty()) continue;
			addOfficeText(content, external_office_text);
			addStyles(content, external_styles_text);
			continue;			
		}		
		if((posTO=key.find("\" TypeOut=\"copy"))!=string::npos){
			addOfficeText(content, office_text);
			continue;			
		}
		if((posTO=key.find("\" TypeOut=\"delete_table"))!=string::npos){
			deleteTable(content, key.substr(0, posTO));
			continue;			
		}		
/*		if((posTO=key.find("\" TypeOut=\"refreshTextFields"))!=string::npos){
			refreshTextFields(content, office_text);
			continue;			
		}		*/
	
		replaceField(content, key, value);
	}
}

void getTemplateFromXml(const char** templateFileName, const string xmlStr){
	string::size_type pos = xmlStr.find("<FILL DocFileName=\"");
	if(pos==string::npos)
		return;
	pos+=19; //strlen("<FILL DocFileName=\"")
	string::size_type pos2 = xmlStr.find("\">",pos);
	if(pos2==string::npos)
		return;	
	if(pos==pos2)
		return;
	*templateFileName = strdup((path2copyfile_templates+xmlStr.substr(pos,pos2-pos)).c_str());
}

int main(int argc, char** argv){
	if(argc<1+2){
		fprintf(stderr,"usage: program xmlfile template [path2copyfile_templates [output file]]\n");
		return 1;		
	}
	const char* pXmlFile = argv[1];
	const char* pTemplateFile = argv[2];
	if(argc>=1+3)
		path2copyfile_templates = argv[3];
	char *output_file_name = (argc>=1+4)?argv[4]:NULL;
	bool f_open_output_file = (output_file_name == NULL);

//Считывание файла	
	struct stat sstat;
	if(::stat(pXmlFile, &sstat)!=0){
		fprintf(stderr,"stat %s error\n", pXmlFile);
		return 1;
	}	
	FILE*fi = fopen(pXmlFile,"rb");
	if(!fi){
		fprintf(stderr,"fopen %s error\n", pXmlFile);
		return 1;
	}
	char *buf = (char*)malloc(sstat.st_size*(1+3));
	if(!buf){
		fprintf(stderr,"malloc error\n");
		return 1;
	}
	char *outbuf	= buf+sstat.st_size;
	size_t inbytesleft = fread(buf,1,sstat.st_size,fi);
	if(inbytesleft!=sstat.st_size){
		fprintf(stderr,"stat %s error\n", pXmlFile);
		return 1;
	}	
	fclose(fi);
//Смена кодировки	
	iconv_t iconv_handle = iconv_open ("UTF-8", "CP1251");
	if(iconv_handle==(iconv_t)(-1)){
		fprintf(stderr,"iconv_open error\n");
		return 1;
	}
	char *inbuf=buf;
	char *p_outbuf=outbuf;
	size_t outbytesleft=sstat.st_size*3;
	size_t r = iconv (iconv_handle,
              &inbuf, &inbytesleft,
              &p_outbuf, &outbytesleft);
	if(r==(size_t)(-1)){
		fprintf(stderr,"iconv error\n");
		return 1;
	}	
	iconv_close(iconv_handle);	
	
	getTemplateFromXml(&pTemplateFile, outbuf);
	
//Упаковка в новый файл на базе шаблона
{
	struct zip *zs;
	int i, idx, err;
	char errstr[1024];
    if ((zs=zip_open(pTemplateFile, 0, &err)) == NULL) {
		zip_error_to_str(errstr, sizeof(errstr), err, errno);
		fprintf(stderr, "cannot open  %s: %s\n",pTemplateFile,errstr);
		return 1;
    }	
    if ((idx=zip_name_locate(zs, "content.xml", ZIP_FL_NODIR|ZIP_FL_NOCASE)) != -1) {
		struct zip_file *zf;
	//Чтение содержимого		
		static string content="";
		if ((zf=zip_fopen_index(zs, idx, 0)) == NULL) {
			fprintf(stderr, "cannot open file %i in archive: %s\n",idx,zip_strerror(zs));
			return 1;
		}
		int n;
		char buf[8192];
		while ((n=zip_fread(zf, buf, sizeof(buf))) > 0) {
			//~ for(int i=0;i<n;i++)
				//~ cout << buf[i];
			for(int i=0;i<n;i++)
				content+=buf[i];
		}		
		if (n < 0) {
			fprintf(stderr, "error reading file %i in archive: %s\n",idx,zip_file_strerror(zf));
			zip_fclose(zf);
			return 1;
		}
		zip_fclose(zf);
		parseAndReplace(&content, outbuf);
		//~ cout <<content;
		struct zip_source *source=zip_source_buffer(zs, content.c_str(), content.length(), 0);		
		if(source==NULL){
			fprintf(stderr, "error at zip_source_buffer\n");
			return 1;
		}
		int r;
		if((r = zip_replace(zs,idx,source))<0){
			zip_error_to_str(errstr, sizeof(errstr), err, errno);			
			fprintf(stderr, "cannot replace %i %s\n",r,errstr);
			return 1;
		}
		//zip_source_free(source);
	}else{
		fprintf(stderr, "zip_name_locate error\n");
	}
	int r;	
	if(!output_file_name){
		if ((output_file_name=(char *)malloc(strlen(pTemplateFile)+30)) == NULL) {
			fprintf(stderr, "malloc error\n");	
			return 1;
		}
		time_t rawtime = time (NULL);
		struct tm * timeinfo = localtime ( &rawtime );	
	
		const char* lastSlash = strrchr(pTemplateFile, SLASH);
		strcpy(output_file_name, (lastSlash)?(lastSlash+1):pTemplateFile);
		char *p_point=strrchr(output_file_name,'.');
		if(!p_point){
			free(output_file_name);
			fprintf(stderr, "template file name must contain point\n");
			return 1;
		}
		char ext[3+1];
		for(int i=0;i<sizeof(ext);i++)
			ext[i]=*(p_point+1+i);
    	sprintf(p_point, "%i_%i_%i__%i_%i_%i.%s", timeinfo->tm_year+1900,timeinfo->tm_mon+1,	timeinfo->tm_mday,timeinfo->tm_hour,timeinfo->tm_min,timeinfo->tm_sec, ext);
	}
	if((r = zip_close_into_new_file(zs, output_file_name))<0)
		fprintf(stderr, "zip_close error %i %s\n", r, output_file_name);
#ifdef _WIN32
//Открытие созданного файла
	if(f_open_output_file){
		ShellExecute(NULL, "open", output_file_name, NULL, NULL, SW_SHOWNORMAL);
	}
#endif
	free(output_file_name);
}
	free(buf);
	return 0;
}
