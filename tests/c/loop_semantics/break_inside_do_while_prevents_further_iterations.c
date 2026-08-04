// vybe-test: c/loop_semantics/break_inside_do_while_prevents_further_iterations
// origin: languages/c/tests/c/test_loop_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"body\n"};
int __n = 1, __i = 0;
int x = 0; do { { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "body");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } while (++x < 3); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

