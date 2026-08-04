// vybe-test: c/lang_pointer_arithmetic_steps/pointer_plus_zero_is_identity
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int a[3]={4,5,6}; int *p=a+1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *(p+0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

