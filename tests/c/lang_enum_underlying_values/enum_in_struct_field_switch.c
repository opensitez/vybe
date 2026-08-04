// vybe-test: c/lang_enum_underlying_values/enum_in_struct_field_switch
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum Mode { IDLE, RUN }; struct Job { enum Mode mode; };
int main() {
const char *__w[] = {"r\n"};
int __n = 1, __i = 0;
struct Job j = {RUN}; switch(j.mode){case IDLE: { char __t[512]; snprintf(__t, sizeof(__t), "i\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case RUN: { char __t[512]; snprintf(__t, sizeof(__t), "r\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

