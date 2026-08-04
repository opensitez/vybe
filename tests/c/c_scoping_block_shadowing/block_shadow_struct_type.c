// vybe-test: c/c_scoping_block_shadowing/block_shadow_struct_type
// origin: languages/c/tests/c/test_c_scoping_block_shadowing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"10"};
int __n = 1, __i = 0;
 int x = 5; { struct S { int x; } s = {10}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

