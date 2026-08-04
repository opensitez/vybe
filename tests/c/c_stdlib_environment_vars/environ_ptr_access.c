// vybe-test: c/c_stdlib_environment_vars/environ_ptr_access
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
 setenv("TEST_ENVIRON", "1", 1); int found = 0; for(char **e = environ; *e; e++) { if((*e)[0] == 'T') found = 1; } { char __t[512]; snprintf(__t, sizeof(__t), "%d", found);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

