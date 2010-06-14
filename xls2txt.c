/*
 *	Copyright (C) 2005-2009 Jan Bobrowski <jb@wizard.ae.krakow.pl>
 *
 *	This program is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU General Public License
 *	version 2 as published by the Free Software Foundation.
 */

/*
 * Based on information from sc.openoffice.org/excelfileformat.pdf
 */

#include "xls2txt.h"
#include <stdio.h>
#include <getopt.h>
#include <time.h>
#include <math.h>

#define TRUNC errx(1, "Truncated  &%d", __LINE__)
#define BADF(T) errx(1, *T""?T"  &%d":"Format error  &%d", __LINE__);

struct g {
	unsigned all:1;
	unsigned sel:1; // -n
	unsigned nofmt:1;
	unsigned titles:1;
	int nr; // sheet number
	int row, col; // current pos
	unsigned top, bottom, left, right;
} g;

struct sst {
	u8 *ptr, *rend;
};

struct fmt {
	unsigned type:8; // 1:num, 3:date, 4:time, 5:date-time
	unsigned arg:8;
};

struct tab {
	void *tab;
	int nelem, aelem, esize;
};

static inline void *tab_ptr(struct tab *tab, unsigned n)
{
	return (char *)tab->tab + n * tab->esize;
}

#define TAB(S,T,N) (*(T*)((char*)(S).tab + (N) * sizeof(T)))

void *tab_alloc(struct tab *tab, unsigned n, const void *dflt)
{
	u8 *p;

	if(n < tab->nelem)
		return tab_ptr(tab, n);
	if(n >= tab->aelem) {
		int sz = n+16 & ~15;
		tab->aelem = sz;
		tab->tab = realloc(tab->tab, sz * tab->esize);
		if(!tab->tab) err(1, "realloc");
	}
	p = tab_ptr(tab, tab->nelem);
	n -= tab->nelem;
	tab->nelem += n + 1;
	if(n) do {
		memcpy(p, dflt, tab->esize);
		p += tab->esize;
	} while(--n);
	return p;
}

struct xls {
	meml_t map;
	u8 *end;
	u8 *shptr;
/*
	00:*	BIFF2
	02:*	BIFF3
	04:*	BIFF4
	08:0500	BIFF5 (and BIFF7)
	08:0600	BIFF8

	08:0000 BIFF5
	08:0200 BIFF2
	08:0300 BIFF3
	08:0400 BIFF4
*/
	enum {BIFF2=2,BIFF3=3,BIFF4=4,BIFF5=5,BIFF8=6} biffv;
	unsigned e1904;

	struct sst *sst;
	unsigned nsst;

	struct tab fmt;
	struct tab xf_ptr;
	struct tab xf_fmt;
};

static struct xls x;

void check_biffv(u8 *p)
{
	int v;
	if(p[0]!=9)
		errx(1, "Format not recognized");
	switch(p[1]) {
	case 0:
biff2:
		errx(1, "BIFF2 not supported");
	case 2: v = BIFF3; goto ok;
	case 4: v = BIFF4; goto ok;
	case 8: break;
	default:
nsupp:
		errx(1, "Format not supported");
	}
	switch(p[5]) {
	case 0: v = BIFF5; break;
	case 2: goto biff2;
	default:
		v = p[5];
		if(v<BIFF3 || v>BIFF8)
			goto nsupp;
	}
ok:
	x.biffv = v;
}

static u8 *print_str(u8 *p, int l)
{
	if(x.biffv < BIFF8) {
		p = print_cp_str(p, l);
	} else {
		u8 f = *p++;
		int a=0;
		if(f&8) {a += 4*g16(p); p += 2;}
		if(f&4) {a += g32(p); p += 4;}
		p = a + print_uni(p, l, f);
	}
	return p;
}

static void print_sst(int n)
{
	u8 *p, *re, f;
	unsigned l;

	if(n<0 || n>=x.nsst)
		BADF("Wrong string index");

	p = x.sst[n].ptr;
	re = x.sst[n].rend;
	l = g16(p); f = p[2]; p += 3;
	p += (f&8 ? 2 : 0) + (f&4 ? 4 : 0);
	for(;;) {
		int s = re - p;
		f &= 1;
		if(l <= s>>f)
			break;

		if(re[0] != 0x3C) // CONTINUE
			BADF("String truncated");

		l -= s>>f;
		if(s&f)
			BADF("String cut at the middle of a char");
		print_uni(p, s>>f, f);

		p = re + 4;
		re = p + g16(re+2);
		f = *p++;
	}
	print_uni(p, l, f);
}

static u8 *read_sst(u8 *p, u8 *re, u8 *fe)
{
	unsigned nsst;

	x.nsst = g32(p+4);
	if(!x.nsst)
		return re;

	x.sst = calloc(x.nsst, sizeof *x.sst);
	if(!x.sst) err(1, "calloc");

	p += 8;
	
	for(nsst = 0;;) {
		unsigned l, a;
		u8 f;

		if(re-p < 3)
			BADF("String table truncated");

		x.sst[nsst].ptr = p;
		x.sst[nsst].rend = re;
		if(++nsst == x.nsst)
			break;

		l = g16(p);
		f = p[2]; p += 3;
		a = 0;
		if(f&8) {a = 4*g16(p); p += 2;}
		if(f&4) {a += g32(p); p += 4;}
//		fmt_assert(p<re);
		for(;;) {
			int s = re - p;
			f &= 1;
			if(s >= l<<f)
				break;
//			fmt_assert(!(s&f));
//			fmt_assert(re < fe);
			l -= s>>f;
			if(re[0] != 0x3C) // CONTINUE
				BADF("String truncated");
			p = re + 4;
			re = p + g16(re+2);
//			fmt_assert(re < fe);
			f = *p++;
		}
		p += l<<f;
		for(;;) {
			int s = re - p;
			if(s > a) break;
			a -= s;

			if(re[0] != 0x3C) // CONTINUE
				BADF("String truncated");
			p = re + 4;
			re = p + g16(re+2);
//			fmt_assert(re < fe);
		}
		p += a;
	}
	return re;
}

static const struct fmt default_fmt;
static const u8 *null_ptr;

static void xls_init_struc()
{
	static u8 t[] = {0,0x10,0x12,0x10,0x12,0x10,0x10,0x12,0x12,0x12,0x14,
		0x22,0,0,0x30,0x30,0x30,0x30,0x40,0x40,0x40,0x40,0x50,0,0,0,
		0,0,0,0,0,0,0,0,0,0,0,0x10,0x10,0x12,0x12,0x10,0x10,0x12,0x12,
		0x40,0x40,0x40,0x21};
	struct fmt *tab;
	int i;

	x.fmt.esize = sizeof default_fmt;
	x.fmt.nelem = 0;
	x.xf_ptr.esize = sizeof null_ptr;
	x.xf_ptr.nelem = 0;
	x.xf_fmt.esize = sizeof null_ptr;
	x.xf_fmt.nelem = 0;
	x.e1904 = 0;

	tab_alloc(&x.fmt, elemof(t)-1, &default_fmt);
	tab = x.fmt.tab;
	for(i=0;i<elemof(t);i++)
		tab[i].type=t[i]>>4, tab[i].arg=t[i]&0xF;
}

static void getstr(u16 *d, u8 *p, int l)
{
	int v = 0;
	if(x.biffv>=BIFF8) v = *p++ & 1;
	// XXX
	if(v)
		while(--l>=0) d[l] = g16(p+2*l);
	else
		while(--l>=0) d[l] = p[l];
}

static void parse_fmt(struct fmt *f, u16 *p, int l)
{
	u16 *e = p + l;
	u16 *q, *d;

	f->type = 0;
	f->arg = 0;

	if(e==p) return;
	q = p;
	while(*q=='[') {
		do
			if(++q==e) return;
		while(*q!=']');
		if(++q==e) return;
	}
	if(*p=='Y'||*p=='M'||*p=='D') {
		f->type = 5;
		return;
	}
	if(*p=='h'||*p=='m') {
		f->type = 4;
		return;
	}

	p = q;
	d = 0;
	for(;;) {
		if(*q=='.') {
			d = q;
			break;
		}
		if(*q>=128 || !strchr("0#?, ", *q) || ++q==e)
			break;
	}
	if(!d) {
		if(p!=q && (q==e || *q!='/')) {
//			f->arg = 0;
			f->type = 1;
		}
		return;
	}
	while(++q<e)
		if(*q!='0' && *q!='#')
			break;

	f->arg = q-d-1;
	f->type = 1;
}

static void
set_fmt(u8 *p)
{
	u8 *q;
	int n, l;
	struct fmt *fmt;
	u16 t[128];

	q = p+1;
	if (x.biffv >= BIFF4) {
		q += 2;
	}
	n = x.biffv < BIFF5 ? x.fmt.nelem : g16(p);
	l = q[-1];
	if(x.biffv >= BIFF8) {
		l = g16(p+2), q++;
	}

	if (l > elemof(t)) {
		return;
	}

	getstr(t, q, l);
	fmt = (struct fmt*)tab_alloc(&x.fmt, n, &default_fmt);
	parse_fmt(fmt, t, l);
}

static const struct fmt *fmt_from_xf(int xf)
{
	const struct fmt *fmt = &default_fmt;
	int n, st, ua, org_xf;
	u8 *p;

	if(xf >= x.xf_ptr.nelem) {
bad_xf:
		warnx("Strange XF index %u -- ignored", xf);
		return fmt;
	}

	org_xf = xf;

again:
	p = TAB(x.xf_ptr, u8*, xf);
	if(!p) goto bad_xf;

	if(x.biffv < BIFF5) { /* 0x02 */
		n = p[1];
		st = 2;
		ua = x.biffv < BIFF4 ? 3 : 5;
	} else { /* 0xE0 */
		n = g16(p+2);
		st = 4;
		ua = x.biffv < BIFF8 ? 7 : 9;
	}
	st = p[st];
	ua = p[ua];

	if(!((st ^ ua) & 4)) { /* format not present */
		if(!(st & 4) || xf!=org_xf) { /* not a style or loop */
			p += x.biffv!=BIFF4 ? 4 : 2;
			xf = g16(p) >> 4;
			if(xf!=org_xf && xf < x.xf_ptr.nelem)
				goto again;
		}
	} else if(n < x.fmt.nelem)
		fmt = &TAB(x.fmt, struct fmt, n);

	*(const struct fmt**)tab_alloc(&x.xf_fmt, org_xf, &null_ptr) = fmt;
	return fmt;
}

static void print_time(int m, int f, double v);

static void print_fmt(unsigned xf, double v)
{
	const struct fmt *f;

	if(g.nofmt) {
		printf("%f", v);
		return;
	}

	if(xf < x.xf_fmt.nelem) {
		f = TAB(x.xf_fmt, struct fmt*, xf);
		if(f) goto have_fmt;
	}
	f = fmt_from_xf(xf);
have_fmt:

	switch(f->type) {
	case 0:
		if(ceil(v)==v) {
			printf("%.f", v);
			break;
		}
	default:
		printf("%f", v);
		break;
	case 1: printf("%.*f", f->arg, v); break;
	case 2: printf("%.*E", f->arg, v); break;
	case 3:
	case 4:
	case 5: print_time(f->type-2, f->arg, v); break;
	}
}

static void print_time(int m, int f, double v)
{
	int d;
	time_t t;
	struct tm *tm;

	d = v; v -= d;
	if(x.e1904)
		d += 4*365;
	else if(d<=60)
		d++;
	d -= 25569;

	t = d*24*60*60 + (unsigned)(v*24*60*60);
	tm = gmtime(&t);
	if(!tm) {
		printf("#BAD"); // XXX
		return;
	}
	if(m==3 && !f && !v)
		m = 1;
	if(m&1) {
		printf("%04u-%02u-%02u", tm->tm_year+1900, tm->tm_mon+1, tm->tm_mday);
		if(m==1) return;
		printf(" ");
	}
	printf("%2u:%02u:%02u", tm->tm_hour, tm->tm_min, tm->tm_sec);
}

static void print_rk(unsigned xf, u32 rk)
{
	double v;
	if(rk & 2)
		v = (s32)rk>>2;
	else
		v = ieee754((u64)(rk&~3) << 32);
	if(rk & 1)
		v /= 100;
	print_fmt(xf, v);
}

struct rr {
	int o, l, id;
};

#define GETRR(P) \
	if(4 > x.map.len-rr.o) TRUNC; \
	rr.l = g16(x.map.ptr+rr.o+2); \
	rr.id = x.map.ptr[rr.o]; \
	rr.o += 4; \
	if(rr.l > x.map.len-rr.o) TRUNC; \
	(P) = x.map.ptr + rr.o; \
	rr.o += rr.l;

#define EXPLEN(L) if(rr.l < (L)) errx(1, "Record too short  &%d", __LINE__);

static int skip_substream(int o)
{
	struct rr rr;
	int d=1;
	rr.o = o;
	for(;;) {
		u8 *p, sv;
		GETRR(p)
		sv = p[-3];
		switch(rr.id) {
		case 0x09:
			if(sv<0x10) d++;
			break;
		case 0x0A:
			if(!sv && !--d)
				return rr.o;
		}
	}
	TRUNC;
}

static int read_init_rr(int o)
{
	struct rr rr;
	int sh, nr;
	u8 *p;

	xls_init_struc();
	rr.o = o;
	nr = g.nr; sh = 0;

	for(;;) {
		GETRR(p)

		switch(rr.id) {
		case 0x42: // CODEPAGE
			set_codepage(g16(p));
			break;
		case 0xFC: // SST
			rr.o = read_sst(p, x.map.ptr+rr.o, x.end) - x.map.ptr;
			break;
		case 0x1E: // FORMAT
			set_fmt(p);
			break;
		case 0x43:
		case 0xE0: // XF
			*(u8**)tab_alloc(&x.xf_ptr, x.xf_ptr.nelem, &null_ptr) = p;
			break;
		case 0x04: // LABEL
		case 0x03: // NUMBER
		case 0x06: // FORMULA
		case 0x07: // STRING
		case 0x7E: // RK
			return sh;
		case 0x09: // BOF
			if (p[-3]>=0x10) {
				break;
			}
			rr.o = skip_substream(rr.o);
			break;
		case 0x0A: // EOF
			if (p[-3]) {
				break;
			}
			return sh;
		case 0x85: // SHEET
			if(!nr--) {
				sh = p - 4 - x.map.ptr;
			}
			break;
		case 0x22: // DATEMODE
			x.e1904 = p[0];
			break;
		}
	}
}

int to_cell(int r, int c)
{
	if(r < g.top || r > g.bottom) {
		g.row = r;
		return 0;
	}
	if(g.row < g.top)
		g.row = g.top;
	if(g.row < r) {
		g.col = 0;
		do {
			putchar('\n');
			g.row++;
		} while(g.row < r);
	}
	if(c < g.left || c > g.right) {
		g.col = c;
		return 0;
	}
	if(g.col < g.left)
		g.col = g.left;
	while(g.col < c) {
		putchar('\t');
		g.col++;
	}
	return 1;
}

static int to_cell_p(u8 *p) {return to_cell(g16(p), g16(p+2));}

static inline int to_nx_cell() {return to_cell(g.row, g.col+1);}

void print_sheet(int o, u8 *name, int nr)
{
	struct rr rr;
	u8 pvrec;

	if(g.titles) {
		if(nr) putchar('\f');
		if(name) print_str(name+1, *name);
		putchar('\n');
	}

	rr.o = o;
	g.col = g.row = 0;
	pvrec = 0;

	for(;;) {
		u8 *p;

		GETRR(p)
		if(rr.id == 0x0A && !p[-3]) // EOF
			break;

		switch(rr.id) {
		case 0x09: // BOF
			if(p[-3]>=0x10) break;
			rr.o = skip_substream(rr.o);
			break;
		case 0x04: // LABEL
			if(to_cell_p(p))
				print_str(p+8, g16(p+6));
			break;
		case 0xFD: // LABELSST
			if(to_cell_p(p))
				print_sst(g16(p+6));
			break;
		case 0x7E: // RK
			if(to_cell_p(p))
				print_rk(g16(p+4), g32(p+6));
			break;
		case 0xBD: { // MULRK
				u8 *q = p + rr.l - 11;
				int f = to_cell_p(p);
				for(;;) {
					unsigned xf = g16(p+4);
					p += 6;
					if(f) print_rk(xf, g32(p));
					if(p>=q) break;
					f = to_nx_cell();
				}
			} break;
		case 0x03: // NUMBER
			if(!to_cell_p(p)) {
				break;
			}
number:
			print_fmt(g16(p+4), ieee754(g64(p+6)));
			break;
		case 0x06: // FORMULA
			if(!to_cell_p(p)) {
				pvrec = 0;
				break;
			}
			if (g16(p+6+6) != 0xFFFF) {
				pvrec = 0;
				goto number;
			}
			break;
		case 0x07: // STRING
			if (pvrec==0x06) {
				print_str(p+2, g16(p));
			}
			break;
		case 0xD6: // RSTRING
			if (to_cell_p(p)) {
				print_str(p+8, g16(p+6));
			}
			break;
		}
		pvrec = rr.id;

		if(g.row > g.bottom)
			break;
	}
	putchar('\n');
}

void print_xls()
{
	struct rr rr;
	int done;
	u8 *p;

	done = 0;
	rr.o = 0;
	GETRR(p)

	switch(g16(p+2)) {
	case 0x10: // single sheet
		if(g.nr) goto not_found;
		read_init_rr(rr.o);
		print_sheet(rr.o, 0, 0);
		return;
	case 0x100: goto workbook;
	case 5: goto globals;
	default:
		BADF("Bad content");
	}

/* BIFF5+ */
globals:
	rr.o = read_init_rr(rr.o);
	if(!rr.o)
		goto not_found;
	for(;;) {
		u32 o;
		GETRR(p)
		if(rr.id != 0x85) // SHEET
			break;
		o = rr.o;
		rr.o = g32(p);
		if(rr.o >= x.map.len)
			TRUNC;
		if(rr.o <= p-x.map.ptr)
			BADF( );
		if(p[4]==0) {
			u8 *q;
			GETRR(q)
			if(rr.id != 0x09) BADF( );
			print_sheet(rr.o, p+6, done++);
			if(!g.all) break;
		} else if(g.sel)
			goto not_found;
		rr.o = o;
	}
	return;

/* BIFF4W */
workbook:
	for(;;) {
		GETRR(p)
		switch(rr.id) {
			u32 o;
		case 0x42: // CODEPAGE
			EXPLEN(2)
			set_codepage(g16(p));
			break;
		case 0x8E: // SHEETOFFSET
			EXPLEN(4)
			o = g32(p);
			if(o >= x.map.len) TRUNC;
			rr.o = o;
			goto found;
		case 0x0A: // EOF
			if(p[-3]) break;
			BADF( );
		}
	}
found:
	GETRR(p)
	if(rr.id != 0x8F) // SHEETHDR
		goto not_found;
	EXPLEN(5)
	for(;;) {
		u32 o = g32(p);
		if(o >= x.map.len-rr.o)
			TRUNC;
		o += rr.o;
		if(!g.nr--) {
			u8 *name = p+4;
			GETRR(p)
			if(rr.id != 0x09) // BOF
				BADF( )
			if(g16(p+2) == 0x10) {
				read_init_rr(rr.o);
				print_sheet(rr.o, name, done++);
				if(!g.all)
					break;
			} else if(g.sel)
				errx(1, "Not a sheet");
			g.nr = 0;
		}
		rr.o = o;
		GETRR(p)
		if(rr.id != 0x8F) {
			if(!done)
				goto not_found;
			break;
		}
	}
	return;

not_found:
		errx(1, "No such sheet");
}

void list_xls()
{
	struct rr rr;
	u8 *p;
	int nr;

	rr.o = 0;
	GETRR(p)
	if(rr.id != 0x09) // BOF
		BADF( );
	switch(g16(p+2)) {
	case 0x10:
		printf("Single sheet\n");
		return;
	case 5:
	case 0x100:
		break;
	default:
		printf("Unknown contents\n");
		return;
	}

	nr = 0;
	for(;;) {
		GETRR(p)
		switch(rr.id) {
			u8 *q;
			char *k;
		case 0x0A: // EOF
			if(p[-3]) break;
			return;
		case 0x09: // BOF
			if(p[-3]>=0x10) break;
			rr.o = skip_substream(rr.o);
			break;
		case 0x42: // CODEPAGE
			set_codepage(g16(p));
			break;
		case 0x85: // SHEET
			k = "sheet";
			q = p;
			if(x.biffv > BIFF4) {
				switch(q[5]) {
				case 0: break;
				case 2: k="chart"; break;
				case 6: k="vbasic"; break;
				default: k=""; break;
				}
				q += 6;
			}
			printf("%2u. %-8s ", nr++, k);
			print_str(q+1, q[0]);
			putchar('\n');
			break;
		}
	}
}

static char *parse_cell(char *s, unsigned *r, unsigned *c)
{
	unsigned a = *s - 'A';
	if(a < 26) {
		unsigned v = a;
		for(;;) {
			a = *++s - 'A';
			if(a >= 26) break;
			v = 26*v + a;
		}
		*c = v;
	}
	a += 'A' - '0';
	if(a < 10) {
		unsigned v = a;
		for(;;) {
			a = *++s - '0';
			if(a >= 10) break;
			v = 10*v + a;
		}
		*r = v - 1;
	}
	return s;
}

void parse_range(char *s)
{
	s = parse_cell(s, &g.top, &g.left);
	if(!*s) return;
	if(*s==':') {
		s = parse_cell(s+1, &g.bottom, &g.right);
		if(!*s) return;
	}
	errx(1, "unexpected char '%c' in cell range", *s);
}

int main(int argc, char *argv[])
{
	char o=0;

	for(;;) switch(getopt(argc, argv, "n:AlC:a12P:fhV?-")) {
		int n;
	case -1: goto endopt;
	case 'n': g.sel=1; g.nr = atoi(optarg); break;
	case 'A': g.sel=0; g.all=1; g.titles=1; break;
	case 'l': o = 'l'; break;
	case 'C':
		n = find_charset(optarg);
		if(n<0) warnx("%s: Unknown charset", optarg);
		set_charset(n);
		break;
	case 'a': set_charset(1); break;
	case '1': set_charset(2); break;
	case '2': set_charset(3); break;
	case 'P':
		n = atoi(optarg);
		if(n) set_codepage(n);
		break;
	case 'f': g.nofmt = 1; break;
	case '?':
		if(optopt!='?') break;
	case '-':
	case 'h':
	case 'V':
#define _STR(T) #T
#define STR(T) _STR(T)
		printf("xls2txt " STR(VERSION) " / "
			"Copyright 2009 Jan Bobrowski / GPL\n");
		goto usage;
	}
endopt:

	g.right = g.bottom = 0xFFFF;
	switch(argc-optind) {
	default:
usage:
		printf(
			"usage: xls2txt [-C cs] [-n sheetnum|-A] [-f] file.xls [X:X]\n"
			"       xls2txt [-C cs] -l file.xls\n"
			" X:X\tcell range (eg. A1:C5, D2:E)\n"
			" -l\tlist sheets\n"
			" -n num\tselect sheet\n"
			" -A\tall sheets (\\f separated)\n"
			" -C cs\toutput charset (utf8 asc iso1 iso2), utf8 is default\n"
			" -f\tdon't try to format numbers\n"
			" -a\tascii output (same as -C asc)\n"
		);
		return 1;
	case 1: break;
	case 2:	parse_range(argv[argc-1]);
		break;
	}

	ole_open(argv[optind]);
	x.map = get_workbook();
	x.end = x.map.ptr + x.map.len;
	check_biffv(x.map.ptr);
	if(o)
		list_xls();
	else
		print_xls();

	return 0;
}
