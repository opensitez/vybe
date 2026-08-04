// vybe-test: c/lang_array_decay_parameters/sizeof_param_does_not_see_caller_bound
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int words(int a[100]){ return (int)(sizeof a); }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
int a[3]={0}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", words(a));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

