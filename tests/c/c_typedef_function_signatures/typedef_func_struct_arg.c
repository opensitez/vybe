// vybe-test: c/c_typedef_function_signatures/typedef_func_struct_arg
// origin: languages/c/tests/c/test_c_typedef_function_signatures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int a; }; typedef int (*F)(struct S); int get(struct S s) { return s.a; } int main() {const char *__w[] = {"88"};
int __n = 1, __i = 0;
 F f = get; struct S s = {88}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", f(s));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

