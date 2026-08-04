// vybe-test: c/preprocessor_advanced/token_paste_operator
// origin: languages/c/tests/c/test_preprocessor_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define CONCAT(a, b) a##b
int main() {const char *__w[] = {"42\n"};
int __n = 1, __i = 0;

    int xy = 42;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", CONCAT(x, y));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

