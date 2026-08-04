// vybe-test: c/lang_array_decay_parameters/callee_mutates_through_decayed_param
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void set0(int a[]){ a[0]=99; }
int main() {
const char *__w[] = {"99\n"};
int __n = 1, __i = 0;
int a[2]={1,2}; set0(a); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

