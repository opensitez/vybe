// vybe-test: c/control_flow/early_return
// origin: languages/c/tests/c/test_control_flow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int sign(int n) {
    if (n > 0) return 1;
    if (n < 0) return -1;
    return 0;
}
int main() {const char *__w[] = {"1\n", "-1\n", "0\n"};
int __n = 3, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sign(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sign(-3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sign(0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

