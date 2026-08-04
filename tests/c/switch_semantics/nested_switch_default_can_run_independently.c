// vybe-test: c/switch_semantics/nested_switch_default_can_run_independently
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"inner-default\n"};
int __n = 1, __i = 0;
int a = 1; int b = 9; switch (a) { case 1: switch (b) { default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "inner-default");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "outer-default");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

