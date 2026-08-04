// vybe-test: c/function_pointers/function_pointer_can_point_to_function_returning_double
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
double half(double x) { return x / 2.0; }
int main() {
const char *__w[] = {"4.50\n"};
int __n = 1, __i = 0;
double (*fp)(double) = half;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.2f\n", fp(9.0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

