// vybe-test: c/lang_preprocessor_conditional_code/nested_if_both_true
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define A
#define B
#if defined(A)
  #if defined(B)
    #define V 9
  #else
    #define V 1
  #endif
#else
  #define V 0
#endif
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", V);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

