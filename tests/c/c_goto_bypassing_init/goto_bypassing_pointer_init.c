// vybe-test: c/c_goto_bypassing_init/goto_bypassing_pointer_init
// origin: languages/c/tests/c/test_c_goto_bypassing_init.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"E"};
int __n = 1, __i = 0;
 int x = 1; goto L; int *p = &x; L: { char __t[512]; snprintf(__t, sizeof(__t), "E");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

