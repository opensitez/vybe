// vybe-test: c/lang_preprocessor_conditional_code/elif_third_branch_matches
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define TIER 3
#if TIER==1
#define SCORE 10
#elif TIER==2
#define SCORE 20
#elif TIER==3
#define SCORE 30
#else
#define SCORE 0
#endif
int main() {
const char *__w[] = {"30\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", SCORE);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

