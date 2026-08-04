// vybe-test: c/lang_vla_stack_arrays/vla_index_from_end_negative_offset
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"40\n"};
int __n = 1, __i = 0;
int n=4; int a[n]; a[0]=10;a[1]=20;a[2]=30;a[3]=40; int *end=&a[n]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", end[-1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

