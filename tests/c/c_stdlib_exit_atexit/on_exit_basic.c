// vybe-test: c/c_stdlib_exit_atexit/on_exit_basic
// vybe-test-exit: 42
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"42 hello"};
static int __n = 1, __i = 0;
#define _BSD_SOURCE
#include <stdlib.h>
#include "on_exit_compat.h"
void func(int status, void *arg) {
  char __t[512]; snprintf(__t, sizeof(__t), "%d %s", status, (char*)arg);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); }
  __i++;
  if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
}
int main() {
  on_exit(func, "hello");
  exit(42);
  return 0;
}
