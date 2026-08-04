// vybe-test: c/lang_vla_stack_arrays/vla_passed_to_function_by_pointer_decay
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int sum_n(int *p, int n){ int s=0; for(int i=0;i<n;i++) s+=p[i]; return s; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int n=3; int a[n]; a[0]=1;a[1]=2;a[2]=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum_n(a,n));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

