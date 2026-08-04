// vybe-test: c/lang_pointer_arithmetic_steps/prefix_decrement_then_postfix_increment
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"6 7\n"};
int __n = 1, __i = 0;
int a[3]={5,6,7}; int *p=&a[2]; int v=*--p; p++; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", v, *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

