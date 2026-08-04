// vybe-test: c/compound_assignment/compound_assignment_on_double_keeps_fraction
// origin: languages/c/tests/c/test_compound_assignment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
double x = 1.5;
int main() {
const char *__w[] = {"1.75\n"};
int __n = 1, __i = 0;
x += 0.25;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.2f\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

