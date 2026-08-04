// vybe-test: c/lang_pointer_comparison_order/same_array_greater_equal_high_index
// origin: languages/c/tests/c/test_lang_pointer_comparison_order.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
int a[4]={0}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", &a[3]>=&a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

