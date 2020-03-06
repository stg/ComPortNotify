windres -i resource.rc resource.o
gcc main.cpp serial.c -lgdi32 -lsetupapi resource.o -mwindows -o bin/cpnotify
del resource.o
strip bin/cpnotify.exe