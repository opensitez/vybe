#include "case_helper.h"
#include <ctype.h>
#include <stddef.h>

char *strlwr(const char *s) {
    static char bufs[4][1024];
    static int idx = 0;
    char *buf = bufs[idx++ & 3];
    size_t i = 0;
    while (s[i] && i < 1023) {
        buf[i] = (char)tolower((unsigned char)s[i]);
        i++;
    }
    buf[i] = '\0';
    return buf;
}

char *strupr(const char *s) {
    static char bufs[4][1024];
    static int idx = 0;
    char *buf = bufs[idx++ & 3];
    size_t i = 0;
    while (s[i] && i < 1023) {
        buf[i] = (char)toupper((unsigned char)s[i]);
        i++;
    }
    buf[i] = '\0';
    return buf;
}
