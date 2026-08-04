// vybe-test: c/switch_semantics/switch_default_can_appear_before_final_case
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"three\n"};
int __n = 1, __i = 0;
int x = 3; switch (x) { default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "other");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case 3: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "three");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

