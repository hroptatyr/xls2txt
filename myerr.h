/*
 * fake up a quick myerr.h: 
 *
 * void err(int eval, const char *fmt, ...); 
 * void errx(int eval, const char *fmt, ...);
 * void warnx(const char *fmt, ...);
 */
#include <errno.h>
#include <unistd.h>
#include <stdio.h>

#define err(eval, fmt, ...)	{ 					\
    (void)fprintf(stderr, "xls2txt: "fmt": ", ##__VA_ARGS__);		\
    (void)fprintf(stderr, "%s\n", strerror(errno));			\
    exit(eval); }

#define errx(eval, fmt, ...)	{ 					\
    (void)fprintf(stderr, "xls2txt: "); 				\
    (void)fprintf(stderr, fmt"\n", ##__VA_ARGS__);			\
    exit(eval); }

#define warnx(fmt, ...) 						\
    (void)fprintf(stderr, "xls2txt: " fmt "\n", ##__VA_ARGS__) 
