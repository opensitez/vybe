// vybe-test: c/c_stdio_ungetc_pushback/ungetc_then_rewind
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"a"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetc_rew.txt", "w+"); fputs("abc", f); rewind(f); ungetc('X', f); rewind(f); { char __t[512]; snprintf(__t, sizeof(__t), "%c", fgetc(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

