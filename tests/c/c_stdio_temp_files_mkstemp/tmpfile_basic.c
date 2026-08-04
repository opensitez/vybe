// vybe-test: c/c_stdio_temp_files_mkstemp/tmpfile_basic
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"hello"};
int __n = 1, __i = 0;
 FILE *f = tmpfile(); if (!f) return 1; fputs("hello", f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

