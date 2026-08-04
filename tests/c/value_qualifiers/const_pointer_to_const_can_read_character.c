// vybe-test: c/value_qualifiers/const_pointer_to_const_can_read_character
// origin: languages/c/tests/c/test_value_qualifiers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
const char *const text = "hi";
int main() {
const char *__w[] = {"i\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", text[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

