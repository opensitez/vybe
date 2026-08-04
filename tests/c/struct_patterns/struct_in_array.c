// vybe-test: c/struct_patterns/struct_in_array
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct { int x; int y; } Point;
int main() {
const char *__w[] = {"2 4\n"};
int __n = 1, __i = 0;
Point pts[3] = {{0,0},{1,1},{2,4}};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", pts[2].x, pts[2].y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

