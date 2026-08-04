// vybe-test: c/control_flow_advanced/continue_in_for_loop_skips_post_continue_statements_only
// origin: languages/c/tests/c/test_control_flow_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n", "tail\n", "2\n", "tail\n"};
int __n = 4, __i = 0;
for (int i = 0; i < 3; i++) { if (i == 1) continue; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "tail");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

