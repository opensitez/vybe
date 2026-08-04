// vybe-test: c/c_ternary_struct_return/ternary_struct_pass_to_function
// origin: languages/c/tests/c/test_c_ternary_struct_return.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"8"};
static int __n = 1, __i = 0;
struct S { int a; }; void f(struct S s) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { struct S s1={8}, s2={9}; f(1 ? s1 : s2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

