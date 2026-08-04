// vybe-test: c/c_switch_nested/switch_nested_continue_in_loop
// origin: languages/c/tests/c/test_c_switch_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 for(int i=0; i<2; i++) { switch(i) { case 0: switch(1) { case 1: continue; } break; case 1: { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

