// vybe-test: c/lang_preprocessor_conditional_code/nested_conditional_sum
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define A 1
#if A
  #define X 10
#else
  #define X 0
#endif
#define B 1
#if B
  #define Y 22
#else
  #define Y 0
#endif
int main() {
const char *__w[] = {"32\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", X+Y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

