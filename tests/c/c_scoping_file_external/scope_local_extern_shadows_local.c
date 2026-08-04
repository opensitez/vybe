// vybe-test: c/c_scoping_file_external/scope_local_extern_shadows_local
// origin: languages/c/tests/c/test_c_scoping_file_external.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int g = 100; int main() {const char *__w[] = {"100"};
int __n = 1, __i = 0;
 int g = 1; { extern int g; { char __t[512]; snprintf(__t, sizeof(__t), "%d", g);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

