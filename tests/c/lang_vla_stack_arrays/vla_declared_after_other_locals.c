// vybe-test: c/lang_vla_stack_arrays/vla_declared_after_other_locals
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int x=1; int n=3; int a[n]; a[0]=x+4; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

