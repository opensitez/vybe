// vybe-test: c/lang_control_scope/static_local_retains
// origin: languages/c/tests/c/test_lang_control_scope.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int bump(){ static int c=0; c++; return c; }
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", bump(), bump());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

