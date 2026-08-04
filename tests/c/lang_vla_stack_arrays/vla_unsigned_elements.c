// vybe-test: c/lang_vla_stack_arrays/vla_unsigned_elements
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"11\n"};
int __n = 1, __i = 0;
int n=2; unsigned a[n]; a[0]=5;a[1]=6; { char __t[512]; snprintf(__t, sizeof(__t), "%u\n", a[0]+a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

