// vybe-test: c/lang_array_decay_parameters/two_d_param_write_visible_in_caller
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void set(int a[][2]){ a[1][1]=77; }
int main() {
const char *__w[] = {"77\n"};
int __n = 1, __i = 0;
int m[2][2]={{0}}; set(m); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", m[1][1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

