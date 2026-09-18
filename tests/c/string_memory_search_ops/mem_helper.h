#ifndef MEM_HELPER_H
#define MEM_HELPER_H

#include <stddef.h>

void *my_memmem(const void *haystack, size_t haystacklen, const void *needle, size_t needlelen);
void *my_memrchr(const void *s, int c, size_t n);

#define memmem my_memmem
#define memrchr my_memrchr

#endif
