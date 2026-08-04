// vybe-test: c/operators_logical/logical_and_allows_right_side_when_left_true
// origin: languages/c/tests/c/test_operators_logical.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int x = 0;
int main() {
const char *__w[] = {"ok\n", "1\n"};
int __n = 2, __i = 0;
if (1 && ++x) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

