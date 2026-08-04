// vybe-test: c/complex_initializers/global_struct_array_init
// origin: languages/c/tests/c/test_complex_initializers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pt { int x; int y; };
struct Pt points[3] = {{1,2},{3,4},{5,6}};
int main() {
const char *__w[] = {"3 6\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", points[1].x, points[2].y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

