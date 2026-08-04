// vybe-test: c/string_escapes/escape_sequence_can_be_used_in_character_comparison
// origin: languages/c/tests/c/test_string_escapes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"match\n"};
int __n = 1, __i = 0;
if ('\n' == 10) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "match");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "bad");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

