// vybe-test: c/bitfields/bitfield_overflow_wraps
// origin: languages/c/tests/c/test_bitfields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct B { unsigned int val : 2; };
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
struct B b;
b.val = 7;
{ char __t[512]; snprintf(__t, sizeof(__t), "%u\n", b.val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

