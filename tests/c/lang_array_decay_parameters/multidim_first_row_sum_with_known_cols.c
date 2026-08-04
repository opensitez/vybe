// vybe-test: c/lang_array_decay_parameters/multidim_first_row_sum_with_known_cols
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int row0_sum(int a[][4]){ return a[0][0]+a[0][1]+a[0][2]+a[0][3]; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int m[2][4]={{1,2,3,4},{0,0,0,0}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", row0_sum(m));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

