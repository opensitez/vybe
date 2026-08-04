// vybe-test: c/lang_preprocessor_breadth/conditional_compilation_selects_code
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define USE_A 1
#if USE_A
#define VAL 7
#else
#define VAL 0
#endif
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", VAL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

