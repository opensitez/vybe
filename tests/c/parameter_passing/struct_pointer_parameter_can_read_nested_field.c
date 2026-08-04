// vybe-test: c/parameter_passing/struct_pointer_parameter_can_read_nested_field
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Point { int x; int y; }; int get_y(struct Point *point) { return point->y; }
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
struct Point point = {3, 4}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", get_y(&point));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

