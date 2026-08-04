// vybe-test: c/c_ternary_struct_return/ternary_struct_nested_structs
// origin: languages/c/tests/c/test_c_ternary_struct_return.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int a; }; struct Outer { struct Inner i; }; int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 struct Outer o1={{1}}, o2={{2}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (0 ? o1 : o2).i.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

