// vybe-test: c/c_stdlib_exit_atexit/on_exit_mixed_with_atexit
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"21"};
static int __n = 1, __i = 0;
#define _BSD_SOURCE
#include <stdlib.h>
void f1() { { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
void f2(int s, void *a) { { char __t[512]; snprintf(__t, sizeof(__t), "2");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() { atexit(f1); on_exit(f2, NULL); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

