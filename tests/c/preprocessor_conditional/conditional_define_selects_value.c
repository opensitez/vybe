// vybe-test: c/preprocessor_conditional/conditional_define_selects_value
// origin: languages/c/tests/c/test_preprocessor_conditional.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define RELEASE
#ifdef RELEASE
#define LOG_LEVEL 0
#else
#define LOG_LEVEL 3
#endif
int main() {const char *__w[] = {"0\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", LOG_LEVEL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

