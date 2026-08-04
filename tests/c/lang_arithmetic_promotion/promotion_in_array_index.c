// vybe-test: c/lang_arithmetic_promotion/promotion_in_array_index
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
int a[10]={0}; char i=3; a[i]=42; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

