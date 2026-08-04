// vybe-test: c/switch_semantics/switch_can_fall_through_into_default
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"one\n", "tail\n"};
int __n = 2, __i = 0;
int x = 1; switch (x) { case 1: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "one");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "tail");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

