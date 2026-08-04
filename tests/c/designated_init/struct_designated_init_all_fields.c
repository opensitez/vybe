// vybe-test: c/designated_init/struct_designated_init_all_fields
// origin: languages/c/tests/c/test_designated_init.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Color { int r; int g; int b; };
int main() {
const char *__w[] = {"64 128 255\n"};
int __n = 1, __i = 0;
struct Color c = {.b=255, .g=128, .r=64};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", c.r, c.g, c.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

