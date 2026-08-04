// vybe-test: c/lang_array_decay_parameters/callee_reads_second_via_decayed_param
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int second(int a[]){ return a[1]; }
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;
int a[3]={10,20,30}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", second(a));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

