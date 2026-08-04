// vybe-test: c/c_io_posix_open_memstream/posix_open_wmemstream_basic
// origin: languages/c/tests/c/test_c_io_posix_open_memstream.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <wchar.h>
#include <stdlib.h>
#include <stdio.h>
int main() {const char *__w[] = {"4"};
int __n = 1, __i = 0;
 wchar_t *buf; size_t size; FILE *f = open_wmemstream(&buf, &size); if (f) { fwprintf(f, L"wide"); fclose(f); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)size);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(buf); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

