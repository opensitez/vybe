// vybe-test: c/anonymous_struct_union/anonymous_union_in_struct
// origin: languages/c/tests/c/test_anonymous_struct_union.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
struct Value {
    int type;
    union {
        int i;
        float f;
    };
};
int main() {const char *__w[] = {"42\n"};
int __n = 1, __i = 0;

    struct Value v;
    v.type = 0;
    v.i = 42;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", v.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

