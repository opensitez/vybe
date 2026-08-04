// vybe-test: c/preprocessor/macro_can_expand_to_boolean_condition
// origin: languages/c/tests/c/test_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define SHOULD_RUN 1
int main() {
const char *__w[] = {"run\n"};
int __n = 1, __i = 0;
if (SHOULD_RUN) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "run");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "stop");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

