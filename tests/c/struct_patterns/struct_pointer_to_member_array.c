// vybe-test: c/struct_patterns/struct_pointer_to_member_array
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Vec { float v[3]; };
int main() {
const char *__w[] = {"1 3\n"};
int __n = 1, __i = 0;
struct Vec vv = {{1.0f, 2.0f, 3.0f}};
float *p = vv.v;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.0f %.0f\n", p[0], p[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

