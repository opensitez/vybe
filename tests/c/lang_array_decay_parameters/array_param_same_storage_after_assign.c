// vybe-test: c/lang_array_decay_parameters/array_param_same_storage_after_assign
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void bump(int a[]){ a[0]++; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int a[1]={4}; bump(a); bump(a); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

