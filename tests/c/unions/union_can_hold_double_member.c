// vybe-test: c/unions/union_can_hold_double_member
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Value { double d; int i; };
int main() {
const char *__w[] = {"2.5\n"};
int __n = 1, __i = 0;
union Value value; value.d = 2.5; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", value.d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

