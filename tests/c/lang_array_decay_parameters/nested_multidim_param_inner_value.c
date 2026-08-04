// vybe-test: c/lang_array_decay_parameters/nested_multidim_param_inner_value
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int deep(int a[][2][2]){ return a[0][1][0]; }
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int m[1][2][2]={{{1,2},{3,4}}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", deep(m));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

