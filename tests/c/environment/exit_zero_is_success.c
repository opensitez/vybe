// vybe-test: c/environment/exit_zero_is_success
// origin: languages/c/tests/c/test_environment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdlib.h>
int main() {const char *__w[] = {"before\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "before\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    exit(EXIT_SUCCESS);
    { char __t[512]; snprintf(__t, sizeof(__t), "after\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

