// vybe-test: c/c_stdio_ungetc_pushback/ungetc_ftell_behavior
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2 1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetc_ftell.txt", "w+"); fputs("abcdef", f); rewind(f); fgetc(f); fgetc(f); long pos1 = ftell(f); ungetc('X', f); long pos2 = ftell(f); { char __t[512]; snprintf(__t, sizeof(__t), "%ld %ld", pos1, pos2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

