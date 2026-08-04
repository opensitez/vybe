// vybe-test: c/c_switch_case_ranges_gnu/switch_case_ranges_single_value_match
// origin: languages/c/tests/c/test_c_switch_case_ranges_gnu.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"B"};
int __n = 1, __i = 0;
 int x=5; switch(x) { case 1 ... 4: { char __t[512]; snprintf(__t, sizeof(__t), "A");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case 5: { char __t[512]; snprintf(__t, sizeof(__t), "B");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case 6 ... 10: { char __t[512]; snprintf(__t, sizeof(__t), "C");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

