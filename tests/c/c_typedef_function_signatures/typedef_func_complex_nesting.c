// vybe-test: c/c_typedef_function_signatures/typedef_func_complex_nesting
// origin: languages/c/tests/c/test_c_typedef_function_signatures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*F1)(int); typedef F1 (*F2)(int); int f1(int x) { return x; } F1 f2(int y) { return f1; } int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 F2 f = f2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", f(1)(2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

