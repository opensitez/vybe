// vybe-test: c/lang_vla_stack_arrays/vla_subarray_via_pointer_offset
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int n=5; int a[n]; for(int i=0;i<n;i++) a[i]=i+1; int *mid=&a[2]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", mid[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

