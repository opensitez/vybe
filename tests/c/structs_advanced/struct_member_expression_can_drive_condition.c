// vybe-test: c/structs_advanced/struct_member_expression_can_drive_condition
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Flag { int on; };
int main() {
const char *__w[] = {"on\n"};
int __n = 1, __i = 0;
struct Flag flag = {1}; if (flag.on) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "on");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "off");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

