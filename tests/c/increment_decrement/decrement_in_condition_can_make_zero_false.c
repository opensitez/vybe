// vybe-test: c/increment_decrement/decrement_in_condition_can_make_zero_false
// origin: languages/c/tests/c/test_increment_decrement.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int x = 1;
int main() {
const char *__w[] = {"false\n", "0\n"};
int __n = 2, __i = 0;
if (--x) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "true");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "false");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

