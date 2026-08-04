// vybe-test: c/lang_preprocessor_conditional_code/conditional_array_size_via_if
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define BIG 1
#if BIG
#define N 4
#else
#define N 2
#endif
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int a[N]={1,2,3,4}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[N-1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

