CXX      := x86_64-w64-mingw32-g++
CXXFLAGS := -Wall -Wextra -O3
LDFLAGS  := -static -s
LIBS     := -lole32 -luuid

ifeq ($(ENABLE_UTF8),)
CXXFLAGS += -D_UNICODE -DUNICODE -municode
endif


all: mkshortcut.exe shortcutinfo.exe

clean:
	-rm -f mkshortcut.exe shortcutinfo.exe

mkshortcut.exe: mkshortcut.cpp
shortcutinfo.exe: shortcutinfo.cpp

%.exe: %.cpp
	$(CXX) $(CXXFLAGS) $^ -o $@ $(LDFLAGS) $(LIBS)

