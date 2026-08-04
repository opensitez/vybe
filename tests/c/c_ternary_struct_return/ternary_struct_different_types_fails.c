// vybe-test: c/c_ternary_struct_return/ternary_struct_different_types_fails
// origin: languages/c/tests/c/test_c_ternary_struct_return.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S1 { int a; }; struct S2 { int a; }; int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* struct S1 s1; struct S2 s2; 1 ? s1 : s2; // type mismatch */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

