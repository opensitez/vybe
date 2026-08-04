// vybe-test: c/c_stdio_fgetpos_fsetpos/fseek_does_not_clear_error
// origin: languages/c/tests/c/test_c_stdio_fgetpos_fsetpos.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_seek_err.txt", "r"); if (!f) return 0; fputc('X', f); int e1 = ferror(f) != 0; fseek(f, 0, SEEK_SET); int e2 = ferror(f) != 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", e1, e2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

