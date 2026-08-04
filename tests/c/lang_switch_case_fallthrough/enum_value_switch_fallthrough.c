// vybe-test: c/lang_switch_case_fallthrough/enum_value_switch_fallthrough
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E{T1=1,T2=2};
int main() {
const char *__w[] = {"1\n", "2\n"};
int __n = 2, __i = 0;
enum E v=T1; switch(v){ case T1: { char __t[512]; snprintf(__t, sizeof(__t), "1\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } case T2: { char __t[512]; snprintf(__t, sizeof(__t), "2\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

