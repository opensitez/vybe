// vybe-test: c/enum_advanced2/enum_as_array_index
// origin: languages/c/tests/c/test_enum_advanced2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Color { RED=0, GREEN=1, BLUE=2 };
const char *color_names[] = {"red","green","blue"};
int main() {
const char *__w[] = {"green\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", color_names[GREEN]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

