// vybe-test: c/recursion/count_digits_recursive
// origin: languages/c/tests/c/test_recursion.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int digits(int n) { return n < 10 ? 1 : 1 + digits(n / 10); }
int main() {
const char *__w[] = {"1 2 4\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", digits(5), digits(42), digits(1000));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

