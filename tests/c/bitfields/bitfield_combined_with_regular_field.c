// vybe-test: c/bitfields/bitfield_combined_with_regular_field
// origin: languages/c/tests/c/test_bitfields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Mixed { int id; unsigned int flag : 1; unsigned int count : 7; };
int main() {
const char *__w[] = {"42 1 100\n"};
int __n = 1, __i = 0;
struct Mixed m;
m.id = 42; m.flag = 1; m.count = 100;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %u %u\n", m.id, m.flag, m.count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

