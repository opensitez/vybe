// vybe-test: c/c_comma_operator_function_args/comma_operator_arg_evaluation_order
// origin: languages/c/tests/c/test_c_comma_operator_function_args.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"A", "ok"};
static int __n = 2, __i = 0;
void f(int a, int b) { { char __t[512]; snprintf(__t, sizeof(__t), "A");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { int x=0; f((x=1, x), (x=2, x)); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

