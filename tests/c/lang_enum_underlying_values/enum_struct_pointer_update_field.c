// vybe-test: c/lang_enum_underlying_values/enum_struct_pointer_update_field
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum S { OFF, ON }; struct D { enum S s; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct D d = {OFF}; struct D *p = &d; p->s = ON; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", d.s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

