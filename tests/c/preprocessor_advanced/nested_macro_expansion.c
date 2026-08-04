// vybe-test: c/preprocessor_advanced/nested_macro_expansion
// origin: languages/c/tests/c/test_preprocessor_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define DOUBLE(x) ((x) * 2)
#define QUAD(x) DOUBLE(DOUBLE(x))
int main() {const char *__w[] = {"20\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", QUAD(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

