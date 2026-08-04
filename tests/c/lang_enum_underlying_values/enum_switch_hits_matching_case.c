// vybe-test: c/lang_enum_underlying_values/enum_switch_hits_matching_case
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum Color { RED, GREEN, BLUE };
int main() {
const char *__w[] = {"g\n"};
int __n = 1, __i = 0;
enum Color c = GREEN; switch(c){case RED: { char __t[512]; snprintf(__t, sizeof(__t), "r\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case GREEN: { char __t[512]; snprintf(__t, sizeof(__t), "g\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "x\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

