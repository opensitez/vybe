// vybe-test: c/control_flow/switch_default
// origin: languages/c/tests/c/test_control_flow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"other\n"};
int __n = 1, __i = 0;

    int x = 99;
    switch (x) {
        case 1: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "one");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
        default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "other");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

