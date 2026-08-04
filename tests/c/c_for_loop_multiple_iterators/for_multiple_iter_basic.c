// vybe-test: c/c_for_loop_multiple_iterators/for_multiple_iter_basic
// origin: languages/c/tests/c/test_c_for_loop_multiple_iterators.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"6"};
int __n = 1, __i = 0;
 int i, j, sum=0; for(i=0, j=0; i<3; i++, j++) sum += i+j; { char __t[512]; snprintf(__t, sizeof(__t), "%d", sum);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

