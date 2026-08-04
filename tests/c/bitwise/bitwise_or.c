// vybe-test: c/bitwise/bitwise_or
// origin: languages/c/tests/c/test_bitwise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"14\n"};
int __n = 1, __i = 0;

    int a = 0b1100;
    int b = 0b1010;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a | b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

