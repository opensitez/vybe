// vybe-test: c/c_scoping_file_external/scope_file_forward_func
// origin: languages/c/tests/c/test_c_scoping_file_external.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f(void); int main() {const char *__w[] = {"60"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", f());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; } int f() { return 60; }

