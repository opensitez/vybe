// vybe-test: c/c_stdlib_exit_atexit/atexit_limits
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
void func() { }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int i, res = 0; for(i=0; i<32; i++) { res |= atexit(func); } { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

