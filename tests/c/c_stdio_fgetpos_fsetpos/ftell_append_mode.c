// vybe-test: c/c_stdio_fgetpos_fsetpos/ftell_append_mode
// origin: languages/c/tests/c/test_c_stdio_fgetpos_fsetpos.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_append_tell.txt", "w"); fputs("123", f); fclose(f); f = fopen("test_append_tell.txt", "a"); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", ftell(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

