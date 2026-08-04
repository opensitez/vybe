// vybe-test: c/conditionals_advanced/if_inside_else_branch_can_still_match
// origin: languages/c/tests/c/test_conditionals_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"else-if\n"};
int __n = 1, __i = 0;
int x = 0;
if (x) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "if");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else if (!x) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "else-if");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

