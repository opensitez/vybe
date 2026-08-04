// vybe-test: c/casting/cast_zero_to_pointer_can_compare_with_null
// origin: languages/c/tests/c/test_casting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int *p = (int *)0;
int main() {
const char *__w[] = {"null\n"};
int __n = 1, __i = 0;
if (p == NULL) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "null");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "bad");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

