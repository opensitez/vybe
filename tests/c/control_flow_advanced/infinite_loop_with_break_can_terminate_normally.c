// vybe-test: c/control_flow_advanced/infinite_loop_with_break_can_terminate_normally
// origin: languages/c/tests/c/test_control_flow_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n", "1\n"};
int __n = 2, __i = 0;
int x = 0;
for (;;) { { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (x == 1) break; x++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

