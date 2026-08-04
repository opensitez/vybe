// vybe-test: c/enums_advanced/enum_value_can_drive_switch_case
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum State { OFF, ON };
int main() {
const char *__w[] = {"off\n"};
int __n = 1, __i = 0;
enum State state = OFF; switch (state) { case OFF: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "off");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case ON: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "on");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

