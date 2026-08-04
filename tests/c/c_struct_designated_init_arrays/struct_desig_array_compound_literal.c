// vybe-test: c/c_struct_designated_init_arrays/struct_desig_array_compound_literal
// origin: languages/c/tests/c/test_c_struct_designated_init_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int arr[2]; }; int main() {const char *__w[] = {"10"};
int __n = 1, __i = 0;
 struct S s; s = (struct S){ .arr[1] = 10 }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

