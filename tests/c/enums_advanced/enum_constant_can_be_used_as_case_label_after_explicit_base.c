// vybe-test: c/enums_advanced/enum_constant_can_be_used_as_case_label_after_explicit_base
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Token { START = 10, END = 20 };
int main() {
const char *__w[] = {"end\n"};
int __n = 1, __i = 0;
int token = END; switch (token) { case START: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "start");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case END: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "end");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

