// vybe-test: c/c_scoping_block_shadowing/block_shadow_parameter
// origin: languages/c/tests/c/test_c_scoping_block_shadowing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f(int x) { { int x = 3; return x; } } int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", f(1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

