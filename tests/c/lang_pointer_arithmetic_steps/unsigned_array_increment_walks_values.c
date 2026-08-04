// vybe-test: c/lang_pointer_arithmetic_steps/unsigned_array_increment_walks_values
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
unsigned a[3]={1u,2u,3u}; unsigned *p=a; unsigned s=0; while(p<a+3){s+=*p; p++;} { char __t[512]; snprintf(__t, sizeof(__t), "%u\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

