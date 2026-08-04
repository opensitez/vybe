// vybe-test: c/lang_pointer_arithmetic_steps/short_array_minus_one_step
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"30\n"};
int __n = 1, __i = 0;
short a[4]={10,20,30,40}; short *p=&a[3]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *(p-1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

