// vybe-test: c/structs_advanced/typedef_struct_value_can_be_declared_and_used
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct { int x; int y; } Point;
int main() {
const char *__w[] = {"8 9\n"};
int __n = 1, __i = 0;
Point point = {8, 9};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", point.x, point.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

