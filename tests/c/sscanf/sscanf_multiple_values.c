// vybe-test: c/sscanf/sscanf_multiple_values
// origin: languages/c/tests/c/test_sscanf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"10 2.5\n"};
int __n = 1, __i = 0;
int a; float b;
sscanf("10 2.5", "%d %f", &a, &b);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %.1f\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

