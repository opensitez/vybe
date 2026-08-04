// vybe-test: c/c_stdio_buffering_setvbuf/fflush_stdout_full_buffered
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"hello"};
int __n = 1, __i = 0;
 setvbuf(stdout, NULL, _IOFBF, 1024); { char __t[512]; snprintf(__t, sizeof(__t), "hello");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fflush(stdout); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

