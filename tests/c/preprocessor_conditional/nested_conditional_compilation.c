// vybe-test: c/preprocessor_conditional/nested_conditional_compilation
// origin: languages/c/tests/c/test_preprocessor_conditional.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define A
#define B
int main() {const char *__w[] = {"both\n"};
int __n = 1, __i = 0;

#ifdef A
  #ifdef B
    { char __t[512]; snprintf(__t, sizeof(__t), "both\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
  #else
    { char __t[512]; snprintf(__t, sizeof(__t), "a only\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
  #endif
#endif
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

