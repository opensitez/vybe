// vybe-test: c/enum_advanced2/enum_in_switch_all_cases
// origin: languages/c/tests/c/test_enum_advanced2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef enum { X, Y, Z } Axis;
int main() {
const char *__w[] = {"y\n"};
int __n = 1, __i = 0;

Axis a = Y;
switch (a) {
    case X: { char __t[512]; snprintf(__t, sizeof(__t), "x\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
    case Y: { char __t[512]; snprintf(__t, sizeof(__t), "y\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
    case Z: { char __t[512]; snprintf(__t, sizeof(__t), "z\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

