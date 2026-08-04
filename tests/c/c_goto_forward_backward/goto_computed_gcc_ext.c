// vybe-test: c/c_goto_forward_backward/goto_computed_gcc_ext
// origin: languages/c/tests/c/test_c_goto_forward_backward.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"C"};
int __n = 1, __i = 0;
 void *ptr = &&L; goto *ptr; { char __t[512]; snprintf(__t, sizeof(__t), "X");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } L: { char __t[512]; snprintf(__t, sizeof(__t), "C");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

