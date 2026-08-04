// vybe-test: c/loop_semantics/for_loop_init_runs_once_before_body
// origin: languages/c/tests/c/test_loop_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n", "2\n"};
int __n = 2, __i = 0;
int x = 0; for (x = 1; x < 3; x++) { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

