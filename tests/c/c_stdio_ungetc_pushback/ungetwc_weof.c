// vybe-test: c/c_stdio_ungetc_pushback/ungetwc_weof
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <wchar.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetwc_weof.txt", "w+"); int res = ungetwc(WEOF, f); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == WEOF);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

