// vybe-test: c/lang_vla_stack_arrays/vla_in_for_loop_repeated_allocation
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int total=0; for(int k=1;k<=3;k++){ int a[k]; for(int i=0;i<k;i++) a[i]=i+1; total+=a[k-1]; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", total);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

