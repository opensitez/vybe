// vybe-test: c/c_struct_designated_init_nested/struct_desig_nested_struct_in_array
// origin: languages/c/tests/c/test_c_struct_designated_init_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int a; }; struct Outer { struct Inner arr[2]; }; int main() {const char *__w[] = {"99"};
int __n = 1, __i = 0;
 struct Outer o = { .arr[1].a = 99 }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", o.arr[1].a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

