// vybe-test: c/functions_advanced/mutual_calls_can_be_ordered_with_prototypes
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int odd(int n);
int even(int n) { return n == 0 ? 1 : odd(n - 1); }
int odd(int n) { return n == 0 ? 0 : even(n - 1); }
int main() {
const char *__w[] = {"1 0\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", even(4), odd(4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

