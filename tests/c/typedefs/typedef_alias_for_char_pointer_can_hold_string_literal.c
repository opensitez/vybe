// vybe-test: c/typedefs/typedef_alias_for_char_pointer_can_hold_string_literal
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef char *Text;
int main() {
const char *__w[] = {"vybe\n"};
int __n = 1, __i = 0;
Text text = "vybe"; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", text);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

