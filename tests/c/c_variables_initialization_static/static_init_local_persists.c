// vybe-test: c/c_variables_initialization_static/static_init_local_persists
// origin: languages/c/tests/c/test_c_variables_initialization_static.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f() { static int a = 0; a++; return a; } int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 f(); { char __t[512]; snprintf(__t, sizeof(__t), "%d", f());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

