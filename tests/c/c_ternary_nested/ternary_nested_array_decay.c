// vybe-test: c/c_ternary_nested/ternary_nested_array_decay
// origin: languages/c/tests/c/test_c_ternary_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 int a1[2]={1,2}, a2[2]={3,4}; int *p = 0 ? a1 : a2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

