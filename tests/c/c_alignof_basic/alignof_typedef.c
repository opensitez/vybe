// vybe-test: c/c_alignof_basic/alignof_typedef
// origin: languages/c/tests/c/test_c_alignof_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef double mydouble; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", _Alignof(mydouble) == _Alignof(double));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

