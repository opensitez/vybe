// vybe-test: c/switch_advanced/switch_default_in_middle
// origin: languages/c/tests/c/test_switch_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"default\n"};
int __n = 1, __i = 0;

int x = 99;
switch (x) {
    case 1: { char __t[512]; snprintf(__t, sizeof(__t), "one\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
    default: { char __t[512]; snprintf(__t, sizeof(__t), "default\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
    case 2: { char __t[512]; snprintf(__t, sizeof(__t), "two\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

