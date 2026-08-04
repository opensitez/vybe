// vybe-test: c/c_ternary_struct_return/ternary_struct_sizeof
// origin: languages/c/tests/c/test_c_ternary_struct_return.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { double d; int i; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct S s1, s2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", sizeof(1 ? s1 : s2) == sizeof(struct S));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

