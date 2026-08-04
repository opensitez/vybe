// vybe-test: c/implicit_conversions/float_to_int_truncates
// origin: languages/c/tests/c/test_implicit_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
float f = 3.9f;
int i = f;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

