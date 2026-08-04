// vybe-test: c/lang_pointer_arithmetic_steps/subtracted_offset_reaches_first
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"11\n"};
int __n = 1, __i = 0;
int a[4]={11,22,33,44}; int *p=&a[3]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *(p-3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

