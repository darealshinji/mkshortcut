CXX      := x86_64-w64-mingw32-g++
CXXFLAGS := -Wall -Wextra -O3 -municode
LDFLAGS  := -static -s
LIBS     := -lole32 -luuid

all: mkshortcut.exe shortcutinfo.exe

clean:
	-rm -f mkshortcut.exe shortcutinfo.exe

mkshortcut.exe: mkshortcut.cpp
shortcutinfo.exe: shortcutinfo.cpp

%.exe: %.cpp
	$(CXX) $(CXXFLAGS) $^ -o $@ $(LDFLAGS) $(LIBS)

