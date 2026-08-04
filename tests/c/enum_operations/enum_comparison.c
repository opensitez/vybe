// vybe-test: c/enum_operations/enum_comparison
// origin: languages/c/tests/c/test_enum_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Dir { NORTH=0, EAST=1, SOUTH=2, WEST=3 };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
enum Dir d = EAST;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", d == EAST ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

