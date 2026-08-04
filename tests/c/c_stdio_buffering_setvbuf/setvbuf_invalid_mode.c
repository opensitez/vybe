// vybe-test: c/c_stdio_buffering_setvbuf/setvbuf_invalid_mode
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_vbuf_inv.txt", "w"); int res = setvbuf(f, NULL, 9999, 1024); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

