// vybe-test: c/switch_semantics/switch_case_can_mutate_outer_variable
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
int x = 1; int y = 0; switch (x) { case 1: y = 7; break; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

