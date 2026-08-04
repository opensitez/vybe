// vybe-test: c/typeof/typeof_preserves_pointer_type
// origin: languages/c/tests/c/test_typeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"99\n"};
int __n = 1, __i = 0;

    int x = 99;
    int *p = &x;
    __typeof__(p) q = p;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *q);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

