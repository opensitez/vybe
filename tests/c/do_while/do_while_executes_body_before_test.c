// vybe-test: c/do_while/do_while_executes_body_before_test
// origin: languages/c/tests/c/test_do_while.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"10\n", "9\n"};
int __n = 2, __i = 0;
int x = 10;
do { { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } x--; } while (x > 8);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

