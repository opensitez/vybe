// vybe-test: c/c_typedef_function_signatures/typedef_func_ptr_args
// origin: languages/c/tests/c/test_c_typedef_function_signatures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*F)(int, int); int add(int a, int b) { return a+b; } int main() {const char *__w[] = {"7"};
int __n = 1, __i = 0;
 F ptr = add; { char __t[512]; snprintf(__t, sizeof(__t), "%d", ptr(3, 4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

