// vybe-test: c/basics/function_call
// origin: languages/c/tests/c/test_basics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int add(int a, int b) {
    return a + b;
}
int main() {const char *__w[] = {"7\n"};
int __n = 1, __i = 0;

    int result = add(3, 4);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", result);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

