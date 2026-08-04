// vybe-test: c/c_vla_multidimensional_dynamic/vla_multi_goto_bypassing_fails
// origin: languages/c/tests/c/test_c_vla_multidimensional_dynamic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* int main() { goto L; int arr[2][2]; L: return 0; } // UB if VLA scope entered by goto, but we test compile */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

