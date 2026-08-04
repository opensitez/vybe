// vybe-test: c/switch_semantics/switch_subject_can_be_char_promoted_to_int
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"A\n"};
int __n = 1, __i = 0;
char c = 'A'; switch (c) { case 65: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "A");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "other");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

