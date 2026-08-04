// vybe-test: c/cover_stdio_h/ferror_clearerr
// origin: languages/c/tests/c/test_cover_stdio_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1 0\n"};
int __n = 1, __i = 0;
FILE *f=fopen("/tmp/vybe_c_err.txt","r"); fgetc(f); int e=ferror(f); clearerr(f); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", e, ferror(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

