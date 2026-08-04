// vybe-test: c/c_pointer_comparison/pointer_comparison_const
// origin: languages/c/tests/c/test_c_pointer_comparison.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int x; const int *p1 = &x; int *p2 = &x; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p1 == p2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

