// vybe-test: c/atexit/atexit_single_handler
// origin: languages/c/tests/c/test_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"main\n", "cleanup\n"};
static int __n = 2, __i = 0;

#include <stdio.h>
#include <stdlib.h>
void cleanup() {
    { char __t[512]; snprintf(__t, sizeof(__t), "cleanup\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    atexit(cleanup);
    { char __t[512]; snprintf(__t, sizeof(__t), "main\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

