windres -i resource.rc resource.o
gcc -Os -ffunction-sections -fdata-sections -fno-exceptions -fno-rtti -flto main.cpp serial.cpp toast.cpp -Wl,--gc-sections -Wl,--as-needed -s -lgdi32 -lsetupapi -lshell32 -lshlwapi -lole32 -lpropsys -luuid -lruntimeobject resource.o -mwindows -o bin/cpnotify
del resource.o
