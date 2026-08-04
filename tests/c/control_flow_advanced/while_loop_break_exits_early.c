// vybe-test: c/control_flow_advanced/while_loop_break_exits_early
// origin: languages/c/tests/c/test_control_flow_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n", "1\n", "2\n"};
int __n = 3, __i = 0;
int i = 0;
while (1) { if (i == 3) break; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

