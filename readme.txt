Download libzip source code from http://www.nih.at/libzip/.

wget http://www.nih.at/libzip/libzip-0.10.1.tar.bz2
tar xf libzip-0.10.1.tar.bz2
cd libzip-0.10.1/
patch -p0 < ../zip_close_into_new_file.diff 
CFLAGS="-O3" ./configure --prefix=`pwd`/../libzip-0.10.1_bin 
make
make install

cd ..
make

cd test
#Unix
LD_LIBRARY_PATH=../libzip-0.10.1_bin/lib ../odtfiller xmlfile.xml template.odt output.odt
#Windows
./odtfiller xmlfile.xml template.odt output.odt
./odtfiller xmlfile.xml template.odt



Insert field into template:
Вставка/Поля/Дополнительно... Переменные/Поле пользователя. Формат - текст.
