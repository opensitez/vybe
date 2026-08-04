// vybe-test: c/c_comma_operator_function_args/comma_operator_in_function_ptr_call
// origin: languages/c/tests/c/test_c_comma_operator_function_args.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"2"};
static int __n = 1, __i = 0;
void f(int a) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { void (*p)(int) = f; p((1, 2)); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

