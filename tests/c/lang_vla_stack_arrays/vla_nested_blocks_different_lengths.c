// vybe-test: c/lang_vla_stack_arrays/vla_nested_blocks_different_lengths
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int outer=0; { int n=2; int a[n]; a[1]=4; outer=a[1]; } { int m=3; int b[m]; b[2]=1; outer+=b[2]; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", outer);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

