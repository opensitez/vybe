// vybe-test: c/lang_switch_case_fallthrough/fallthrough_into_case_with_declaration_block
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"a\n", "3\n"};
int __n = 2, __i = 0;
switch(1){ case 1: { char __t[512]; snprintf(__t, sizeof(__t), "a\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } case 2: { int z=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", z);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

