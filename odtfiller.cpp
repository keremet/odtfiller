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
#endif

#include <string>
using namespace std;

void replaceAll(string *content, const string key, const string value){
	for(string::size_type pos=0;(pos = content->find(key, pos))!=string::npos;){
		content->replace(pos,key.length(), value);
	}
}

void replaceField(string *content, const string key, const string value){
	fprintf(stderr,"|%s|%s|\n", key.c_str(), value.c_str());
	string field = "<text:user-field-get text:name=\""+key+"\">"+key+"</text:user-field-get>";	
	replaceAll(content, field, value);
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

void parseAndReplace(string *content, const string xmlStr){
	string::size_type  office_text_posB = content->find("<office:text");
	if(office_text_posB==string::npos) return;
	office_text_posB = content->find('>',office_text_posB);
	if(office_text_posB==string::npos) return;
	office_text_posB++;
	string::size_type  office_text_posE=content->find("</office:text>",office_text_posB); 
	if(office_text_posE==string::npos) return;
	string office_text = content->substr(office_text_posB,office_text_posE-office_text_posB); //Основное содержимое шаблона

	for(string::size_type pos=0;(pos = xmlStr.find("<Field Name=\"", pos))!=string::npos;){
		string key, value;
		string::size_type pos0=pos+13; //strlen("<Field Name=\"");
		if((pos = xmlStr.find("\">", pos0))!=string::npos){
			key=xmlStr.substr(pos0, pos-pos0);
		}else break;
		pos0=pos+2; //strlen("\">")
		if((pos = xmlStr.find("</Field>", pos0))!=string::npos){
			value=xmlStr.substr(pos0, pos-pos0);
		}else break;
		string::size_type posTO;
		if((posTO=key.find("\" TypeOut=\"code"))!=string::npos){
			key=key.substr(0, posTO);
			replaceAll(&value, "&lt;","<");
			replaceAll(&value, "&gt;",">");
			replaceAll(&value, "&quot;","\"");
			//fprintf(stderr,"TypeOut\n");
			if(key == "openoffice_styles_list"){
				addStyles(content, value);
			}else
				replaceFieldWithParagraph(content, key, value);
			continue;			
		}
		if((posTO=key.find("\" TypeOut=\"copy"))!=string::npos){
			content->insert(content->find("</office:text>"),office_text);//!!
			continue;			
		}		
		
		replaceField(content, key, value);
	}
}

int main(int argc, char** argv){
	if(argc<1+2){
		fprintf(stderr,"usage: program xmlfile template [output file]\n");
		return 1;		
	}
	const char* pXmlFile = argv[1];
	const char* pTemplateFile = argv[2];
	char *output_file_name = (argc>=1+3)?argv[3]:NULL;
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
	const char *inbuf=buf;
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
	
		strcpy(output_file_name, pTemplateFile);
		char *p_point=strrchr(output_file_name,'.');
		if(!p_point){
			free(output_file_name);
			fprintf(stderr, "template file name must contain point\n");
			return 1;
		}
    		sprintf(p_point, "%i_%i_%i__%i_%i_%i.odt", timeinfo->tm_year+1900,timeinfo->tm_mon+1,	timeinfo->tm_mday,timeinfo->tm_hour,timeinfo->tm_min,timeinfo->tm_sec);
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
