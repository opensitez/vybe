// vybe-test: c/lang_switch_case_fallthrough/switch_inside_if_with_fallthrough
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"a\n", "b\n"};
int __n = 2, __i = 0;
int x=1; if(x){ switch(x){ case 1: { char __t[512]; snprintf(__t, sizeof(__t), "a\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } case 2: { char __t[512]; snprintf(__t, sizeof(__t), "b\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

