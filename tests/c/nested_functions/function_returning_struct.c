// vybe-test: c/nested_functions/function_returning_struct
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct { int x; int y; } Point;
Point make_point(int x, int y) { Point p = {x, y}; return p; }
int main() {
const char *__w[] = {"3 4\n"};
int __n = 1, __i = 0;
Point p = make_point(3, 4);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", p.x, p.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

