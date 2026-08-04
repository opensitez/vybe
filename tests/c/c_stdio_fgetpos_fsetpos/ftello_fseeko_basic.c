// vybe-test: c/c_stdio_fgetpos_fsetpos/ftello_fseeko_basic
// origin: languages/c/tests/c/test_c_stdio_fgetpos_fsetpos.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ftello.txt", "w+"); fputs("0123", f); fseeko(f, 2, SEEK_SET); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", (long)ftello(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

