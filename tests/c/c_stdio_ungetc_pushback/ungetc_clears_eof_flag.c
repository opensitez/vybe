// vybe-test: c/c_stdio_ungetc_pushback/ungetc_clears_eof_flag
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 0 A"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetc_ceof.txt", "w+"); fgetc(f); int e1 = feof(f) != 0; ungetc('A', f); int e2 = feof(f) != 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %c", e1, e2, fgetc(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

