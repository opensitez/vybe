// vybe-test: c/conditional_logic/logical_and_short_circuits
// origin: languages/c/tests/c/test_conditional_logic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
int x = 0;
int y = x && (1/x > 0);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

