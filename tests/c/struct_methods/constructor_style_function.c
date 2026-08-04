// vybe-test: c/struct_methods/constructor_style_function
// origin: languages/c/tests/c/test_struct_methods.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

typedef struct { int x; int y; int z; } Vec3;
Vec3 vec3_new(int x, int y, int z) { Vec3 v = {x, y, z}; return v; }
int vec3_dot(Vec3 a, Vec3 b) { return a.x*b.x + a.y*b.y + a.z*b.z; }
int main() {
const char *__w[] = {"32\n"};
int __n = 1, __i = 0;
Vec3 a = vec3_new(1,2,3);
Vec3 b = vec3_new(4,5,6);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", vec3_dot(a,b));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

