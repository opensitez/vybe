// vybe-test: c/lang_semantics_batch/array_parameter_decay
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int sum(int *a, int n){ int t=0; for(int i=0;i<n;i++) t+=a[i]; return t; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int a[]={1,2,3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum(a,3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

