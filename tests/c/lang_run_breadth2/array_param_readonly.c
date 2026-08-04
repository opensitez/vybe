// vybe-test: c/lang_run_breadth2/array_param_readonly
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int sum2(const int *a){return a[0]+a[1];}
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int a[]={2,3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum2(a));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

