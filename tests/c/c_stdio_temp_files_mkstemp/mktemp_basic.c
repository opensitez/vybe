// vybe-test: c/c_stdio_temp_files_mkstemp/mktemp_basic
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#include <stdlib.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 char tmpl[] = "test_mktemp_XXXXXX"; char *p = mktemp(tmpl); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p == tmpl);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

