// vybe-test: c/typedef_patterns/typedef_unsigned_byte
// origin: languages/c/tests/c/test_typedef_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef unsigned char u8;
int main() {
const char *__w[] = {"200\n"};
int __n = 1, __i = 0;
u8 x = 200;
{ char __t[512]; snprintf(__t, sizeof(__t), "%u\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

