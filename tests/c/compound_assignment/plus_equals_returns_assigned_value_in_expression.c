// vybe-test: c/compound_assignment/plus_equals_returns_assigned_value_in_expression
// origin: languages/c/tests/c/test_compound_assignment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int x = 1; int y = 0;
int main() {
const char *__w[] = {"5 5\n"};
int __n = 1, __i = 0;
y = (x += 4);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", x, y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

