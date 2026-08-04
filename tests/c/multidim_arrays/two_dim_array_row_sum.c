// vybe-test: c/multidim_arrays/two_dim_array_row_sum
// origin: languages/c/tests/c/test_multidim_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"15\n"};
int __n = 1, __i = 0;

int m[3][3] = {{1,2,3},{4,5,6},{7,8,9}};
int sum = 0;
for (int j = 0; j < 3; j++) sum += m[1][j];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

