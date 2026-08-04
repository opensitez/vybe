// vybe-test: c/loop_semantics/continue_skips_to_next_iteration_in_while_loop
// origin: languages/c/tests/c/test_loop_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n", "3\n", "4\n"};
int __n = 3, __i = 0;
int i = 0; while (i < 4) { i++; if (i == 2) continue; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

