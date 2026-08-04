// vybe-test: c/lang_preprocessor_macro_expansion/macro_expands_inside_switch_case
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define TAG 2
int main() {
const char *__w[] = {"hit\n"};
int __n = 1, __i = 0;
switch(2){case TAG: { char __t[512]; snprintf(__t, sizeof(__t), "hit\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "miss\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

