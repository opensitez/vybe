// vybe-test: c/c_stdlib_environment_vars/putenv_remove
// origin: languages/c/tests/c/test_c_stdlib_environment_vars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _XOPEN_SOURCE
#include <stdlib.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 char var[] = "RM_VAR=1"; putenv(var); char rm[] = "RM_VAR"; putenv(rm); /* Some implementations remove if no = */ { char __t[512]; snprintf(__t, sizeof(__t), "%d", getenv("RM_VAR") == NULL || getenv("RM_VAR") != NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

