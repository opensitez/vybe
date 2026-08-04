// vybe-test: c/typedef_patterns/typedef_multidim_array
// origin: languages/c/tests/c/test_typedef_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Matrix3x3[3][3];
int main() {
const char *__w[] = {"1 1 1\n"};
int __n = 1, __i = 0;
Matrix3x3 m = {{1,0,0},{0,1,0},{0,0,1}};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", m[0][0], m[1][1], m[2][2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

