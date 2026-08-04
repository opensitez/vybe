// vybe-test: c/recursion/gcd_recursive
// origin: languages/c/tests/c/test_recursion.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int gcd(int a, int b) { return b == 0 ? a : gcd(b, a % b); }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", gcd(48, 18));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

