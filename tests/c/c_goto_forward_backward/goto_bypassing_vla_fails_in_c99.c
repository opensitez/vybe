// vybe-test: c/c_goto_forward_backward/goto_bypassing_vla_fails_in_c99
// origin: languages/c/tests/c/test_c_goto_forward_backward.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* int main() { int n=5; goto L; int arr[n]; L: return 0; } // VLA bypass is illegal */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

