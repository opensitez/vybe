// vybe-test: c/lang_pointer_comparison_order/struct_member_array_element_order
// origin: languages/c/tests/c/test_lang_pointer_comparison_order.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct S{int a[3];};
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct S s={{1,2,3}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", &s.a[0]<&s.a[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

