// vybe-test: c/functions/multiple_params
// origin: languages/c/tests/c/test_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int max(int a, int b) {
    return a > b ? a : b;
}
int main() {const char *__w[] = {"7\n", "9\n"};
int __n = 2, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", max(3, 7));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", max(9, 2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

