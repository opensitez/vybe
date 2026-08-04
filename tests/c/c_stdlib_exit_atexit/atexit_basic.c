// vybe-test: c/c_stdlib_exit_atexit/atexit_basic
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"func2 func1 "};
static int __n = 1, __i = 0;
#include <stdlib.h>
void func1() { { char __t[512]; snprintf(__t, sizeof(__t), "func1 ");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
void func2() { { char __t[512]; snprintf(__t, sizeof(__t), "func2 ");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() { atexit(func1); atexit(func2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

