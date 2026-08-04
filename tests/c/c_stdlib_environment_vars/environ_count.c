// vybe-test: c/c_stdlib_environment_vars/environ_count
// origin: languages/c/tests/c/test_c_stdlib_environment_vars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <stdlib.h>
#include <unistd.h>
extern char **environ;
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int count = 0; while(environ && environ[count]) count++; { char __t[512]; snprintf(__t, sizeof(__t), "%d", count >= 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

