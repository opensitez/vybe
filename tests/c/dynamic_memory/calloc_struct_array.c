// vybe-test: c/dynamic_memory/calloc_struct_array
// origin: languages/c/tests/c/test_dynamic_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Point { int x; int y; };
int main() {
const char *__w[] = {"0 0 5 10\n"};
int __n = 1, __i = 0;

struct Point *pts = (struct Point*)calloc(3, sizeof(struct Point));
pts[1].x = 5; pts[1].y = 10;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", pts[0].x, pts[0].y, pts[1].x, pts[1].y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
free(pts);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

