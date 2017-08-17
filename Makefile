# gmake

NAME = xls2txt
VERSION = 0.15
BINDEST = /usr/local/bin
PKG=$(NAME)-$(VERSION)
FILES = Makefile xls2txt.[ch] ole.c cp.c ummap.[ch] ieee754.c list.h myerr.h

CFLAGS ?= -O2 -g -Wall
LDFLAGS = -lm

xls2txt: xls2txt.o ole.o cp.o ummap.o ieee754.o

xls2txt.o: xls2txt.c xls2txt.h
	$(CC) $(CFLAGS) -DVERSION=$(VERSION) -c $< -o $@

install: xls2txt
	install -s $< $(BINDEST)

clean:
	rm -f xls2txt $(addsuffix .o,$(basename $(filter %.c %.[ch],$(FILES))))

dist:
	ln -s . $(PKG)
	tar czf $(PKG).tar.gz --group=root --owner=root $(addprefix $(PKG)/, $(FILES)); \
	rm $(PKG)

check: xls2txt
	lldb -s dbg -- ./$< -l Workbook1.xls
	lldb -s dbg -- ./$< Workbook1.xls

.PHONY: install clean dist check
