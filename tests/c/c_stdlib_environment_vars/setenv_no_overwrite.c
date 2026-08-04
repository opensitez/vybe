// vybe-test: c/c_stdlib_environment_vars/setenv_no_overwrite
// origin: languages/c/tests/c/test_c_stdlib_environment_vars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdlib.h>
int main() {const char *__w[] = {"123"};
int __n = 1, __i = 0;
 setenv("MY_TEST_VAR", "123", 1); setenv("MY_TEST_VAR", "456", 0); { char __t[512]; snprintf(__t, sizeof(__t), "%s", getenv("MY_TEST_VAR"));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

