#include <sys/types.h>
#include <fcntl.h>
#include <stdlib.h>
#include <string.h>
#include <err.h>

typedef unsigned char u8;
typedef unsigned short u16;
typedef unsigned int u32;
typedef signed int s32;
typedef unsigned long
#ifndef __LP64__
long
#endif
u64;

#ifdef __i386__
#define g16(P) (*(u16*)(P))
#define g32(P) (*(u32*)(P))
#define g64(P) (*(u64*)(P))
#define p16(P,V) (*(u16*)(P)=(V))
#else
static inline u16 g16(void *p) {return ((u8*)p)[0] | ((u8*)p)[1]<<8;}
static inline u32 g32(void *p) {return g16(p) | g16((u8*)p+2)<<16;}
static inline u64 g64(void *p) {return g32(p) | (u64)g32((u8*)p+4)<<32;}
static inline void p16(void *p, u16 v) {((u8*)p)[0]=v; ((u8*)p)[1]=v>>8;}
#endif

#define elemof(T) (sizeof T/sizeof*T)
#define endof(T) (T+elemof(T))

typedef struct {
	u8 *ptr;
	unsigned len;
} meml_t;

double ieee754(u64);

int ole_open(char *name);
meml_t get_workbook();

int find_charset(char *name);
void set_charset(int n);	// output charset
u8 *print_uni(u8 *p, int l, u8 f);
void set_codepage(int n);	// sheet codepage
u8 *print_cp_str(u8 *p, int l);
