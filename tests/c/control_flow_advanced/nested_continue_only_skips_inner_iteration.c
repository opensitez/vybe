// vybe-test: c/control_flow_advanced/nested_continue_only_skips_inner_iteration
// origin: languages/c/tests/c/test_control_flow_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"00\n", "02\n", "10\n", "12\n"};
int __n = 4, __i = 0;
for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue; { char __t[512]; snprintf(__t, sizeof(__t), "%d%d\n", i, j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

