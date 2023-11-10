# This one Makefile works with Microsoft nmake and GNU make.
# They use different conditional syntax, but each can be
# nested and inverted within the other.

all: default

ifdef MAKEDIR: # gmake: false; nmake: unused target
!ifdef MAKEDIR # gmake: not seen; nmake: true

#
# Microsoft nmake
#
RM       = del /Q /F
CXX      = cl.exe
CXXFLAGS = -nologo -W3 -O2 -D_UNICODE -DUNICODE
OUT      = -link -out:

!else # and now the other
else

#
# GNU make
#
RM        = rm -f
CXX      := x86_64-w64-mingw32-g++
CXXFLAGS := -Wall -Wextra -O3
LDFLAGS  := -static -s
LIBS     := -lole32 -luuid
OUT       = -o
ifeq ($(ENABLE_UTF8),)
CXXFLAGS += -D_UNICODE -DUNICODE -municode
endif

endif    # gmake: close condition; nmake: not seen
!endif : # gmake: unused target; nmake close conditional


# default target for both
default: mkshortcut.exe shortcutinfo.exe

clean:
	$(RM) *.exe *.o *.obj

mkshortcut.exe: mkshortcut.cpp
	$(CXX) $(CXXFLAGS) mkshortcut.cpp $(LDFLAGS) $(LIBS) $(OUT)mkshortcut.exe

shortcutinfo.exe: shortcutinfo.cpp
	$(CXX) $(CXXFLAGS) shortcutinfo.cpp $(LDFLAGS) $(LIBS) $(OUT)shortcutinfo.exe

