// vybe-test: c/lang_enum_underlying_values/enum_switch_with_explicit_value_labels
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum Http { OK = 200, NOT_FOUND = 404 };
int main() {
const char *__w[] = {"nf\n"};
int __n = 1, __i = 0;
enum Http h = NOT_FOUND; switch(h){case OK: { char __t[512]; snprintf(__t, sizeof(__t), "ok\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case NOT_FOUND: { char __t[512]; snprintf(__t, sizeof(__t), "nf\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

