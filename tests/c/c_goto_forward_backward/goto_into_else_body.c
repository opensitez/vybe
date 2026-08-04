// vybe-test: c/c_goto_forward_backward/goto_into_else_body
// origin: languages/c/tests/c/test_c_goto_forward_backward.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"Else"};
int __n = 1, __i = 0;
 goto L; if(1) {} else { L: { char __t[512]; snprintf(__t, sizeof(__t), "Else");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

