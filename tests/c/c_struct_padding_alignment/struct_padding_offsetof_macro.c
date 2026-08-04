// vybe-test: c/c_struct_padding_alignment/struct_padding_offsetof_macro
// origin: languages/c/tests/c/test_c_struct_padding_alignment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define my_offsetof(t, m) ((size_t)&(((t*)0)->m))
struct S { char c; int i; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", my_offsetof(struct S, i) > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

