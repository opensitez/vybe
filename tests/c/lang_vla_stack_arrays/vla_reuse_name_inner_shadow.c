// vybe-test: c/lang_vla_stack_arrays/vla_reuse_name_inner_shadow
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"7\n", "1\n"};
int __n = 2, __i = 0;
int n=2; int a[n]; a[0]=1; { int n=3; int a[n]; a[2]=7; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

