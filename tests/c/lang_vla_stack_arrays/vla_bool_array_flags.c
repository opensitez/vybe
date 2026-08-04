// vybe-test: c/lang_vla_stack_arrays/vla_bool_array_flags
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int n=3; _Bool f[n]; f[0]=1;f[1]=0;f[2]=1; int c=0; for(int i=0;i<n;i++) if(f[i]) c++; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

