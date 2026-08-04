// vybe-test: c/lang_vla_stack_arrays/vla_copy_into_fixed_buffer
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int n=3; int src[n]; src[0]=1;src[1]=2;src[2]=3; int dst[3]; for(int i=0;i<n;i++) dst[i]=src[i]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", dst[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

