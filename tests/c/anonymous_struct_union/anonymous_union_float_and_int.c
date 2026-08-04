// vybe-test: c/anonymous_struct_union/anonymous_union_float_and_int
// origin: languages/c/tests/c/test_anonymous_struct_union.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
struct Tagged {
    int tag;
    union { int as_int; float as_float; };
};
int main() {const char *__w[] = {"1.5\n"};
int __n = 1, __i = 0;

    struct Tagged t;
    t.tag = 1;
    t.as_float = 1.5f;
    { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", t.as_float);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

