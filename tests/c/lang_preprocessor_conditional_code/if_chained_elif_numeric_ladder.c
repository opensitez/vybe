// vybe-test: c/lang_preprocessor_conditional_code/if_chained_elif_numeric_ladder
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define LEVEL 4
#if LEVEL==1
#define V 1
#elif LEVEL==2
#define V 2
#elif LEVEL==3
#define V 3
#elif LEVEL==4
#define V 27
#else
#define V 0
#endif
int main() {
const char *__w[] = {"27\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", V);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

