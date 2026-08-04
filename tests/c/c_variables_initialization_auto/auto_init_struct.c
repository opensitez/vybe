// vybe-test: c/c_variables_initialization_auto/auto_init_struct
// origin: languages/c/tests/c/test_c_variables_initialization_auto.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int a; int b; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct S s = {1, 2}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.a + s.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

