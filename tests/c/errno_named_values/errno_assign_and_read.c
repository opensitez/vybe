// vybe-test: c/errno_named_values/errno_assign_and_read
// origin: languages/c/tests/c/test_errno_named_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <errno.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
errno=EINVAL; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", errno==EINVAL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

