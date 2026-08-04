// vybe-test: c/c_stdio_ungetc_pushback/ungetc_then_fread
// origin: languages/c/tests/c/test_c_stdio_ungetc_pushback.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"Zell"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ungetc_fread.txt", "w+"); fputs("hello", f); rewind(f); fgetc(f); ungetc('Z', f); char buf[5] = {0}; fread(buf, 1, 4, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

