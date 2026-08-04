// vybe-test: c/c_io_scanf_assignment_suppression/scanf_suppress_float
// origin: languages/c/tests/c/test_c_io_scanf_assignment_suppression.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 4.56"};
int __n = 1, __i = 0;
 float f; int n = sscanf("1.23 4.56", "%*f %f", &f); { char __t[512]; snprintf(__t, sizeof(__t), "%d %.2f", n, f);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

