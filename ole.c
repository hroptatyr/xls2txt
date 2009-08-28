/*
 *	Copyright (C) 2006 Jan Bobrowski <jb@wizard.ae.krakow.pl>
 *
 *	This program is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU General Public License
 *	version 2 as published by the Free Software Foundation.
 */

/*
 * Based on information from sc.openoffice.org/compdocfileformat.pdf
 */

#include "xls2txt.h"
#include <sys/stat.h>
#include <sys/mman.h>
#include <unistd.h>
#include <string.h>
#include "ummap.h"

#define BADSEC (-5)

struct stream_kind {
	unsigned secsc;
	unsigned secsz;
	u32 maxsec;
	s32 (*sat_get)(struct stream_kind *sk, u32 n);
	u8 *(*sec_ptr)(struct stream_kind *sk, u32 n);
};

struct stream {
	struct stream_kind *kind;
	s32 start;
	s32 c_sec;
	unsigned c_pos;
	u8 *c_ptr;
};

struct ole {
	meml_t map;
	int fd;
	char *name;

	s32 root;
	unsigned sec_tshld;
	struct stream ssat;
	struct stream container;

	s32 msat[109];
	s32 msat_start;
//	s32 msat_size;

	struct stream_kind large_sec;
	struct stream_kind small_sec;
} ole;

#define oleerr(S) errx(1, "%s: %s", ole.name, S);
#define oleerrf(F,A...) errx(1, "%s: " F, ole.name, A);

static meml_t mmap_fd(int fd) {
	struct stat st;
	meml_t m;
	if(fstat(fd, &st)<0) err(1, "fstat");
	m.ptr = mmap(0, st.st_size, PROT_READ, MAP_SHARED, ole.fd, 0);
	if((long)m.ptr==-1) err(1, "mmap");
	m.len = st.st_size;
	return m;
}

static s32 sat_get_lg(struct stream_kind *sk, u32 n);
static u8 *sec_ptr_lg(struct stream_kind *sk, u32 n);
static s32 sat_get_sm(struct stream_kind *sk, u32 n);
static u8 *sec_ptr_sm(struct stream_kind *sk, u32 n);

int ole_open(char *name)
{
	u8 h[0x200];
	int v;

	ole.name = name;
	v = open(name, O_RDONLY);
	if(v<0) err(1, "%s", name);
	ole.fd = v;

	v = read(ole.fd, h, sizeof h);
	if(v<sizeof h) {
		if(v<0) err(1, "%s", name);
		errx(1, "%s: File truncated", name);
	}

	ole.map.ptr = 0;
	if(memcmp(h, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", 8))
		return 0;

	if(g16(h+0x1C) != 0xFFFE) oleerr("Not LE file");

	{
		u8 *s = h+76;
		int i;
		ole.msat_start = g32(h+68);
//		ole.msat_size = g32(h+72);
		for(i=0; i<elemof(ole.msat); i++)
			ole.msat[i] = (s32)g32(s), s+=4;
	}

	ole.map = mmap_fd(ole.fd);
	ole.map.ptr += 512;

	{
		struct stream_kind *sk = &ole.large_sec;
		sk->secsc = g16(h+30);
		sk->secsz = 1<<sk->secsc;
		sk->maxsec = g32(h+44) << ole.large_sec.secsc-2;
		sk->sat_get = sat_get_lg;
		sk->sec_ptr = sec_ptr_lg;
	}

	ole.sec_tshld = g32(h+56);
	{
		struct stream_kind *sk = &ole.small_sec;
		sk->secsc = g16(h+32);
		sk->secsz = 1<<sk->secsc;
		sk->maxsec = g32(h+64) << ole.large_sec.secsc-2;
		sk->sat_get = sat_get_sm;
		sk->sec_ptr = sec_ptr_sm;
	}

	ole.ssat.start = g32(h+60);

	ole.root = g32(h+48);
	if(ole.root < 0)
		oleerr("There's no root stream");

	return 1;
}

static void str_open(struct stream *str, struct stream_kind *sk, s32 start)
{
	str->start = start;
	str->c_sec = start;
	str->c_pos = 0;
	str->kind = sk;
	str->c_ptr = sk->sec_ptr(sk, start);
}

#define SID_OK(K,N) ((u32)(N)<=(K)->maxsec)
#define SID_GET(P,I) ((s32)g32((s32*)(P)+(I)))

static s32 sat_get_lg(struct stream_kind *sk, u32 n)
{
	unsigned m, maxsecidx;
	s32 b;

	maxsecidx = (1 << sk->secsc-2) - 1;
	m = n >> sk->secsc-2; n &= maxsecidx;
	if(m < elemof(ole.msat))
		b = ole.msat[m];
	else {
		u8 *p;
		b = ole.msat_start;
		m -= elemof(ole.msat);
		for(;;) {
			if(!SID_OK(sk, b))
				return BADSEC;
			p = sk->sec_ptr(sk, b);
			if(m < maxsecidx)
				break;
			b = SID_GET(p, maxsecidx);
			m -= maxsecidx;
		}
		b = SID_GET(p, m);
	}
	if(SID_OK(sk, b)) {
		u8 *p = sk->sec_ptr(sk, b);
		return SID_GET(p, n);
	}
	return BADSEC;
}

static int str_seek(struct stream *str, unsigned o);

static u8 *sec_ptr_lg(struct stream_kind *sk, u32 n)
{
	return ole.map.ptr + (n<<sk->secsc);
}

static s32 sat_get_sm(struct stream_kind *sk, u32 n)
{
	int o = str_seek(&ole.ssat, 4*n);
	if(o<0) return BADSEC;
	return g32(ole.ssat.c_ptr + o);
}

static u8 *sec_ptr_sm(struct stream_kind *sk, u32 n)
{
	int o = str_seek(&ole.container, n<<sk->secsc);
	if(o<0) oleerr("small sector not found");
	return ole.container.c_ptr + o;
}

static int str_seek(struct stream *str, unsigned o)
{
	struct stream_kind *sk = str->kind;
	unsigned e = str->c_pos + sk->secsz;
	s32 b = str->c_sec;

	if(o < e) {
		if(o >= str->c_pos)
			goto ret;
		e = sk->secsz;
		b = str->start;
		if(o < e) goto found;
	}
	do {
		b = sk->sat_get(sk, b);
		if(!SID_OK(sk, b)) return -1;
		e += sk->secsz;
	} while(o >= e);

found:
	str->c_sec = b;
	str->c_pos = e - sk->secsz;
	str->c_ptr = sk->sec_ptr(sk, b);
ret:
	return o - str->c_pos;
}

static void open_small_streams()
{
	struct stream_kind *sk = &ole.large_sec;
	u8 *p = sec_ptr_lg(sk, ole.root);

	if(!SID_OK(sk, ole.ssat.start) ||
	 !SID_OK(sk, g32(p+0x74))) oleerr("Small sector storage empty");

	str_open(&ole.container, &ole.large_sec, g32(p+0x74));
	str_open(&ole.ssat, &ole.large_sec, ole.ssat.start);
}

static struct ummap wbk_um;
static struct stream wbk_str;

/* this is executed by the signal handler */
static int str_get_page(struct ummap *um, u8 *d)
{
	struct stream_kind *sk = wbk_str.kind;
	int n, c, l;
	u8 *s;

	n = str_seek(&wbk_str, d - (u8*)um->addr);
	if(n<0) return n;

	sk = wbk_str.kind;
	c = sk->secsz - n;
	s = wbk_str.c_ptr + n;

	n = um_access_page(d);
	if(n<0) return n;

	l = um_page_sz - c;
	if(l <= 0) {
		memcpy(d, s, um_page_sz);
		return 0;
	}
	memcpy(d, s, c);
	d += c;

	for(;;) {
		s32 b = sk->sat_get(sk, wbk_str.c_sec);
		if(!SID_OK(sk, b)) return 0;
		s = sk->sec_ptr(sk, b);
		wbk_str.c_sec = b;
		wbk_str.c_pos += sk->secsz;
		wbk_str.c_ptr = s;

		if(l <= sk->secsz) break;
		l -= sk->secsz;
		memcpy(d, s, sk->secsz);
		d += sk->secsz;
	}
	memcpy(d, s, l);

	return 0;
}

static u8 *find_slot(char *name)
{
	struct stream_kind * const sk = &ole.large_sec;
	s32 b;
	u8 *p, *e;
	u16 l;

	b = ole.root;
	p = sk->sec_ptr(sk, b);
	l = 2*(strlen(name) + 1);
	e = p + sk->secsz;
	for(;;) {
		if(p[0x42]==2 && g16(p+0x40)==l) {
			unsigned i = 0;
			for(;; i++) {
				if(2*i >= l)
					return p; // found
				if(p[2*i] != (u8)name[i] || p[2*i+1])
					break;
			}
		}
		p += 0x80;
		if(p < e) continue;

		b = sk->sat_get(sk, b);
		if(!SID_OK(sk, b)) break;
		p = sk->sec_ptr(sk, b);
		e = p + sk->secsz;
	}
	return 0;
}

meml_t get_workbook()
{
	struct stream_kind *sk;
	u32 len, sid;
	u8 *p;

	if(!ole.map.ptr)
		return mmap_fd(ole.fd);

	p = find_slot("Workbook");
	if(!p) {
		p = find_slot("Book");
		if(!p)
			oleerr("No Workbook found");
	}

	sid = g32(p+0x74);
	len = g32(p+0x78);

	sk = &ole.large_sec;
	if(len < ole.sec_tshld) {
		if(!ole.container.c_ptr)
			open_small_streams();
		sk = &ole.small_sec;
	}

	if(!SID_OK(sk, sid))
		oleerr("Stream is empty");

	str_open(&wbk_str, sk, sid);

	wbk_um.size = len;
	wbk_um.handler = (int(*)(struct ummap*,void*))str_get_page;

	if(um_map(&wbk_um) < 0)
		err(1, "um_map");

	return (meml_t){wbk_um.addr, wbk_um.size};
}
