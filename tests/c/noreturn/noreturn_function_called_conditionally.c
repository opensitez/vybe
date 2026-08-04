// vybe-test: c/noreturn/noreturn_function_called_conditionally
// origin: languages/c/tests/c/test_noreturn.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"alive\n"};
static int __n = 1, __i = 0;

#include <stdio.h>
#include <stdlib.h>
_Noreturn void die(void) {
    { char __t[512]; snprintf(__t, sizeof(__t), "dead\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    exit(0);
}
int main() {
    int x = 1;
    if (x > 100) die();
    { char __t[512]; snprintf(__t, sizeof(__t), "alive\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

