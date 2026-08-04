// vybe-test: c/lang_array_decay_parameters/double_array_param_reads_value
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
double pick(double a[], int i){ return a[i]; }
int main() {
const char *__w[] = {"2.5\n"};
int __n = 1, __i = 0;
double d[2]={1.5,2.5}; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", pick(d,1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

