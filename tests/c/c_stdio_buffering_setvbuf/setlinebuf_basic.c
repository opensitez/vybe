// vybe-test: c/c_stdio_buffering_setvbuf/setlinebuf_basic
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_setlinebuf.txt", "w"); setlinebuf(f); fputs("X", f); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

