// vybe-test: c/lang_enum_underlying_values/enum_typedef_switch_case
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
typedef enum { EAST, WEST } Dir;
int main() {
const char *__w[] = {"w\n"};
int __n = 1, __i = 0;
Dir d = WEST; switch(d){case EAST: { char __t[512]; snprintf(__t, sizeof(__t), "e\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; case WEST: { char __t[512]; snprintf(__t, sizeof(__t), "w\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break;} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

