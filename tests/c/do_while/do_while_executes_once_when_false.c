// vybe-test: c/do_while/do_while_executes_once_when_false
// origin: languages/c/tests/c/test_do_while.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"once\n"};
int __n = 1, __i = 0;
int x = 0;
do { { char __t[512]; snprintf(__t, sizeof(__t), "once\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } x++; } while (x < 1);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

