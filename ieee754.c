#define _ISOC99_SOURCE
#include <math.h>
#include "xls2txt.h"

#ifndef __i386__

double ieee754(u64 v)
{
	int s, e;
	double r;

	s = v>>52;
	v &= 0x000FFFFFFFFFFFFFull;
	e = s & 0x7FF;
	if(!e)
		goto denorm;
	if(e < 0x7FF) {
		v += 0x0010000000000000ull, e--;
denorm:
		r = ldexp(v, e - 0x3FF - 52 + 1);
	} else if(v) {
		r = NAN; s ^= 0x800;
	} else
		r = INFINITY;
	if(s & 0x800)
		r = -r;
	return r;
}

#else

double ieee754(u64 v)
{
	union {
		u64 v;
		double d;
	} u;
	u.v = v;
	return u.d;
}

#endif
