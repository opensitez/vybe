// vybe-test: c/c_io_scanf_scansets_matching/scanf_scanset_include_caret
// origin: languages/c/tests/c/test_c_io_scanf_scansets_matching.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"^"};
int __n = 1, __i = 0;
 char buf[10]; sscanf("^abc", "%[^]a-z]", buf); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

