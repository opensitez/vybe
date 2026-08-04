// vybe-test: c/structs/struct_basic
// origin: languages/c/tests/c/test_structs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
struct Point {
    int x;
    int y;
};
int main() {const char *__w[] = {"3\n", "4\n"};
int __n = 2, __i = 0;

    struct Point p;
    p.x = 3;
    p.y = 4;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

