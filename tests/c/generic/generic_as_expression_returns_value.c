// vybe-test: c/generic/generic_as_expression_returns_value
// origin: languages/c/tests/c/test_generic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define ABS(x) _Generic((x), int: abs((int)(x)), double: fabs((double)(x)))
#include <math.h>
#include <stdlib.h>
int main() {const char *__w[] = {"5\n"};
int __n = 1, __i = 0;

    int n = -5;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ABS(n));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

