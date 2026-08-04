// vybe-test: c/lang_run_breadth2/null_ptr_compare_fn
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int f(int x){return x;}
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
int (*fp)(int)=f; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fp!=0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

