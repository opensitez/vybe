// vybe-test: c/loop_semantics/do_while_can_continue_and_reach_condition_check
// origin: languages/c/tests/c/test_loop_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2\n", "3\n"};
int __n = 2, __i = 0;
int x = 0; do { x++; if (x == 1) continue; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } while (x < 3); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

