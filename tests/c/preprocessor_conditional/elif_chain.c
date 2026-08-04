// vybe-test: c/preprocessor_conditional/elif_chain
// origin: languages/c/tests/c/test_preprocessor_conditional.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define PLATFORM 2
int main() {const char *__w[] = {"platform2\n"};
int __n = 1, __i = 0;

#if PLATFORM == 1
    { char __t[512]; snprintf(__t, sizeof(__t), "platform1\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
#elif PLATFORM == 2
    { char __t[512]; snprintf(__t, sizeof(__t), "platform2\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
#else
    { char __t[512]; snprintf(__t, sizeof(__t), "other\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
#endif
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

