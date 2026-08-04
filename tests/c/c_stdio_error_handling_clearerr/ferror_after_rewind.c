// vybe-test: c/c_stdio_error_handling_clearerr/ferror_after_rewind
// origin: languages/c/tests/c/test_c_stdio_error_handling_clearerr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 0"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ferr_rewind.txt", "r"); if (!f) return 0; fputc('X', f); int err1 = ferror(f) != 0; rewind(f); int err2 = ferror(f) != 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", err1, err2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

