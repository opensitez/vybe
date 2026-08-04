// vybe-test: c/lang_enum_underlying_values/enum_switch_fallthrough_two_cases
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { A, B, C };
int main() {
const char *__w[] = {"ab\n"};
int __n = 1, __i = 0;
enum E e = A; switch(e){case A: case B: { char __t[512]; snprintf(__t, sizeof(__t), "ab\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case C: { char __t[512]; snprintf(__t, sizeof(__t), "c\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

