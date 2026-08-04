// vybe-test: c/environment/exit_runs_atexit_handlers
// origin: languages/c/tests/c/test_environment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"hi\n", "bye\n"};
static int __n = 2, __i = 0;

#include <stdio.h>
#include <stdlib.h>
void goodbye(void) { { char __t[512]; snprintf(__t, sizeof(__t), "bye\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() {
    atexit(goodbye);
    { char __t[512]; snprintf(__t, sizeof(__t), "hi\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    exit(0);
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }

