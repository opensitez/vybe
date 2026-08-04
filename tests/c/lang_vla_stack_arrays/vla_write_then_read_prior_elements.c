// vybe-test: c/lang_vla_stack_arrays/vla_write_then_read_prior_elements
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1 9\n"};
int __n = 1, __i = 0;
int n=3; int a[n]; a[2]=9; a[0]=1; a[1]=2; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a[0], a[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

