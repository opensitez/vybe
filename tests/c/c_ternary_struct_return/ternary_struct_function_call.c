// vybe-test: c/c_ternary_struct_return/ternary_struct_function_call
// origin: languages/c/tests/c/test_c_ternary_struct_return.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int a; }; struct S f1() { return (struct S){5}; } struct S f2() { return (struct S){6}; } int main() {const char *__w[] = {"6"};
int __n = 1, __i = 0;
 struct S res = 0 ? f1() : f2(); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

