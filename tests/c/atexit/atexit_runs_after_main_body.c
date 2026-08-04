// vybe-test: c/atexit/atexit_runs_after_main_body
// origin: languages/c/tests/c/test_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"a\n", "b\n", "done\n"};
static int __n = 3, __i = 0;

#include <stdio.h>
#include <stdlib.h>
void done() { { char __t[512]; snprintf(__t, sizeof(__t), "done\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() {
    atexit(done);
    { char __t[512]; snprintf(__t, sizeof(__t), "a\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "b\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

