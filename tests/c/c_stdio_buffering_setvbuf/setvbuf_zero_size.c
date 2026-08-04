// vybe-test: c/c_stdio_buffering_setvbuf/setvbuf_zero_size
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_vbuf_z.txt", "w"); char buf[1024]; int res = setvbuf(f, buf, _IOFBF, 0); /* size 0 might fail or act weird, test it handles it */ { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 0 || res != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

