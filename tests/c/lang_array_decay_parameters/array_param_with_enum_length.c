// vybe-test: c/lang_array_decay_parameters/array_param_with_enum_length
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum{N=3}; int last(int a[]){ return a[N-1]; }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
int a[3]={1,2,8}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", last(a));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

