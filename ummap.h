/*
 *	Copyright (C) 2005 Jan Bobrowski <jb@wizard.ae.krakow.pl>
 *
 *	This library is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU Lesser General Public
 *	License version 2.1 as published by the Free Software Foundation.
 */

#include "list.h"

#ifndef container_of
#include <stddef.h>
#define container_of(P,T,M) ((T*)((char*)(P)-offsetof(T,M)))
#endif

struct ummap {
	list_t list;
	void *addr;
	int size;
	int (*handler)(struct ummap *, void *);
};

extern unsigned um_page_sc, um_page_sz;

int um_map(struct ummap *um);
void um_unmap(struct ummap *um);
int um_access_page(void *p);
