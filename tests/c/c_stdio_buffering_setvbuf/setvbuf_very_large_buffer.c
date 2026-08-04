// vybe-test: c/c_stdio_buffering_setvbuf/setvbuf_very_large_buffer
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_vbuf_lg.txt", "w"); char *buf = malloc(1024*1024); int res = setvbuf(f, buf, _IOFBF, 1024*1024); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); free(buf); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

