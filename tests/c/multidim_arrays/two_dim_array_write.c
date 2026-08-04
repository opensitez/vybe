// vybe-test: c/multidim_arrays/two_dim_array_write
// origin: languages/c/tests/c/test_multidim_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"10 20 30 40\n"};
int __n = 1, __i = 0;

int m[2][2];
m[0][0] = 10; m[0][1] = 20;
m[1][0] = 30; m[1][1] = 40;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", m[0][0], m[0][1], m[1][0], m[1][1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

