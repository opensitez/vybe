// vybe-test: c/c_compound_literals_nested/compound_literal_in_for_init
// origin: languages/c/tests/c/test_c_compound_literals_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"0", "1"};
int __n = 2, __i = 0;
 for (int *p = (int[]){0}; *p < 2; (*p)++) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

