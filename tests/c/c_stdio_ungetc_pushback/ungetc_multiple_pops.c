// vybe-test: c/c_stdio_ungetc_pushback/ungetc_multiple_pops
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"BA3"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetc_pops.txt", "w+"); fputs("123", f); rewind(f); fgetc(f); fgetc(f); ungetc('A', f); ungetc('B', f); { char __t[512]; snprintf(__t, sizeof(__t), "%c%c%c", fgetc(f), fgetc(f), fgetc(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

