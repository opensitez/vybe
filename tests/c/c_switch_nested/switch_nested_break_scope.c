// vybe-test: c/c_switch_nested/switch_nested_break_scope
// origin: languages/c/tests/c/test_c_switch_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"InnerEnd"};
int __n = 1, __i = 0;
 int x=1, y=2; switch(x) { case 1: switch(y) { case 2: break; } { char __t[512]; snprintf(__t, sizeof(__t), "InnerEnd");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

