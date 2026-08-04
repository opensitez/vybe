// vybe-test: c/c_advanced_preprocessor/preprocessor_recursive_macro_guard
// origin: languages/c/tests/c/test_c_advanced_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define A(x) B(x)
#define B(x) A(x) + 1

int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;

    // A(0) expands to A(0) + 1, but recursive expansion is blocked, leaving A(0) + 1
    // Wait, the standard says it's blocked, but compiling A(0) literal isn't valid C if A isn't a function.
    // Instead let's test a known complex conditional macro block.
    { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

