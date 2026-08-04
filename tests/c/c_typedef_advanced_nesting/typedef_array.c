// vybe-test: c/c_typedef_advanced_nesting/typedef_array
// origin: languages/c/tests/c/test_c_typedef_advanced_nesting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int intarr[3]; int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 intarr a = {1, 2, 3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

