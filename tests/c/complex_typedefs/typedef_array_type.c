// vybe-test: c/complex_typedefs/typedef_array_type
// origin: languages/c/tests/c/test_complex_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int IntArr[4];
int main() {
const char *__w[] = {"1 4\n"};
int __n = 1, __i = 0;
IntArr a = {1, 2, 3, 4};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a[0], a[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

