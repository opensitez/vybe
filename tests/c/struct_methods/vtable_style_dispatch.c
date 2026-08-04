// vybe-test: c/struct_methods/vtable_style_dispatch
// origin: languages/c/tests/c/test_struct_methods.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

typedef struct {
    const char *(*name)(void);
    int (*area)(int, int);
} ShapeOps;
const char *rect_name(void) { return "rect"; }
int rect_area(int w, int h) { return w * h; }
int main() {
const char *__w[] = {"rect 12\n"};
int __n = 1, __i = 0;
ShapeOps ops = {rect_name, rect_area};
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %d\n", ops.name(), ops.area(3,4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

