// vybe-test: c/multidim_arrays/three_dim_array
// origin: languages/c/tests/c/test_multidim_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3 6\n"};
int __n = 1, __i = 0;
int arr[2][2][2] = {{{1,2},{3,4}},{{5,6},{7,8}}};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", arr[0][1][0], arr[1][0][1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

