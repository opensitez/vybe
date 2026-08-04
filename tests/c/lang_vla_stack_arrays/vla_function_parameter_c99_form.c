// vybe-test: c/lang_vla_stack_arrays/vla_function_parameter_c99_form
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int last_elem(int n, int a[n]){ return a[n-1]; }
int main() {
const char *__w[] = {"40\n"};
int __n = 1, __i = 0;
int vals[4]={10,20,30,40}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", last_elem(4, vals));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

