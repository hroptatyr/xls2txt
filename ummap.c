/*
 *	Copyright (C) 2005 Jan Bobrowski <jb@wizard.ae.krakow.pl>
 *
 *	This library is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU Lesser General Public
 *	License version 2.1 as published by the Free Software Foundation.
 */

/* These procedures allow the user to employ virtual memory to map
 * arbitrary data to memory. The data can then be computed on-demand
 * instead of preparing it on start.
 */

#include <sys/types.h>
#include <sys/mman.h>
#include <unistd.h>
#include <signal.h>
#include <string.h> // ffs
#include <err.h>
#include "ummap.h"

unsigned um_page_sz, um_page_sc;

static void um_sig(int n, siginfo_t *i, void *c);
static struct sigaction um_sa;
static LIST(maps);

static void um_init()
{
	um_page_sz = getpagesize();
	um_page_sc = ffs(um_page_sz) - 1;

	um_sa.sa_sigaction = um_sig;
	um_sa.sa_flags = SA_SIGINFO|SA_RESETHAND;
}

static void um_sig(int n, siginfo_t *i, void *c)
{
	struct ummap *um;
	unsigned long o;

	if(i->si_code == SEGV_ACCERR
#ifdef __OpenBSD__  // XXX others too?
//#if #system(bsd)
		|| i->si_code == SEGV_MAPERR
#endif
	) {
		list_t *l;
		for(l=maps.next; l!=&maps; l=l->next) {
			um = list_item(l, struct ummap, list);
			o = (char*)i->si_addr - (char*)um->addr;
			if(o < um->size)
				goto found;
		}
	}
	return;

found:
	if(um->handler(um, (char*)um->addr + (o & -um_page_sz)) >= 0)
	 	sigaction(SIGSEGV, &um_sa, 0);
}

int um_access_page(void *p)
{
#if 0
	return (int)mmap(
		p, um_page_sz,
		PROT_READ|PROT_WRITE,
		MAP_PRIVATE|MAP_ANON|MAP_FIXED,
		-1, 0) == -1 ? -1 : 0;
#else
	return mprotect(p, um_page_sz, PROT_READ|PROT_WRITE);
#endif
}

int um_map(struct ummap *um)
{
	void *p;
	int v;

	if(!um_page_sz)
		um_init();

	p = mmap(0, um->size, PROT_NONE, MAP_PRIVATE|MAP_ANON, -1, 0);
	if(p==MAP_FAILED)
		return -1;
	um->addr = p;
	v = sigaction(SIGSEGV, &um_sa, 0);
	if(v>=0) list_add(&maps, &um->list);
	else munmap(p, um->size);
	return v;
}

void um_unmap(struct ummap *um)
{
	munmap(um->addr, um->size);
}
