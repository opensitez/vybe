// vybe-test: c/c_stdio_fgetpos_fsetpos/fsetpos_invalid
// origin: languages/c/tests/c/test_c_stdio_fgetpos_fsetpos.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_inv2.txt", "w"); fclose(f); fpos_t pos; { char __t[512]; snprintf(__t, sizeof(__t), "%d", fsetpos(f, &pos) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

