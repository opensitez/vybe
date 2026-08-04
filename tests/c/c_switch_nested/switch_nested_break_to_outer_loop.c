// vybe-test: c/c_switch_nested/switch_nested_break_to_outer_loop
// origin: languages/c/tests/c/test_c_switch_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"End"};
int __n = 1, __i = 0;
 while(1) { switch(1) { case 1: switch(2) { case 2: break; } break; } { char __t[512]; snprintf(__t, sizeof(__t), "End");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

