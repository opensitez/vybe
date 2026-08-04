// vybe-test: c/structs/struct_init
// origin: languages/c/tests/c/test_structs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
struct Rect {
    int w;
    int h;
};
int main() {const char *__w[] = {"50\n"};
int __n = 1, __i = 0;

    struct Rect r = {10, 5};
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", r.w * r.h);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

