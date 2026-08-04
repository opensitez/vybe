// vybe-test: c/complex_typedefs/typedef_struct_self_named
// origin: languages/c/tests/c/test_complex_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct Vec2 { float x; float y; } Vec2;
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
Vec2 v = {1.0f, 2.0f};
{ char __t[512]; snprintf(__t, sizeof(__t), "%.0f %.0f\n", v.x, v.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

