// vybe-test: c/switch_advanced/switch_fallthrough_intentional
// origin: languages/c/tests/c/test_switch_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"one two"};
int __n = 1, __i = 0;

int x = 1;
switch (x) {
    case 1:
        { char __t[512]; snprintf(__t, sizeof(__t), "one ");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    case 2:
        { char __t[512]; snprintf(__t, sizeof(__t), "two\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        break;
    case 3:
        { char __t[512]; snprintf(__t, sizeof(__t), "three\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

