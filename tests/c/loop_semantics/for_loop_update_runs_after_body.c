// vybe-test: c/loop_semantics/for_loop_update_runs_after_body
// origin: languages/c/tests/c/test_loop_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"b0\n", "u\n", "b1\n", "u\n"};
int __n = 4, __i = 0;
for (int i = 0; i < 2; { char __t[512]; snprintf(__t, sizeof(__t), "u\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }, i++) { char __t[512]; snprintf(__t, sizeof(__t), "b%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

