// vybe-test: c/preprocessor/macro_can_expand_inside_switch_case_value
// origin: languages/c/tests/c/test_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define MATCH 2
int main() {
const char *__w[] = {"hit\n"};
int __n = 1, __i = 0;
int x = 2;
switch (x) { case MATCH: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "hit");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "miss");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

