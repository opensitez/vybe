// vybe-test: c/lang_vla_stack_arrays/vla_function_modifies_caller_buffer
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void dbl(int n, int a[n]){ for(int i=0;i<n;i++) a[i]*=2; }
int main() {
const char *__w[] = {"6 8\n"};
int __n = 1, __i = 0;
int n=2; int a[n]; a[0]=3;a[1]=4; dbl(n,a); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a[0], a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

