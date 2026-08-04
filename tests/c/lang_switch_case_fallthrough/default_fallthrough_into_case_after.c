// vybe-test: c/lang_switch_case_fallthrough/default_fallthrough_into_case_after
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"d\n", "100\n"};
int __n = 2, __i = 0;
switch(99){ default: { char __t[512]; snprintf(__t, sizeof(__t), "d\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } case 100: { char __t[512]; snprintf(__t, sizeof(__t), "100\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

