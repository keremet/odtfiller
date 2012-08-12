APP=odtfiller
OBJ=$(APP).o 
CXXFLAGS=-O3 -Ilibzip-0.10.1_bin/include -Ilibzip-0.10.1_bin/lib/libzip/include
all:$(APP)
$(APP): $(OBJ)
	g++ -o$(APP) $(OBJ) -lzip -Llibzip-0.10.1_bin/lib -lz
clean:
	rm -f $(APP) $(OBJ)
