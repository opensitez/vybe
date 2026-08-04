// vybe-test: c/c_stdlib_environment_vars/setenv_large_value
// origin: languages/c/tests/c/test_c_stdlib_environment_vars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdlib.h>
#include <string.h>
int main() {const char *__w[] = {"9999"};
int __n = 1, __i = 0;
 char large[10000]; memset(large, 'x', 9999); large[9999] = '\0'; setenv("LARGE_VAR", large, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)strlen(getenv("LARGE_VAR")));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

