// vybe-test: c/c_stdio_buffering_setvbuf/fflush_null
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"out"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "out");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fflush(NULL); /* flushes all open streams */ if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

