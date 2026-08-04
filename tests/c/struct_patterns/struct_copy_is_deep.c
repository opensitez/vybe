// vybe-test: c/struct_patterns/struct_copy_is_deep
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Val { int x; int y; };
int main() {
const char *__w[] = {"1 99\n"};
int __n = 1, __i = 0;
struct Val a = {1, 2};
struct Val b = a;
b.x = 99;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a.x, b.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

