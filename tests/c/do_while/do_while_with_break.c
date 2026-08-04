// vybe-test: c/do_while/do_while_with_break
// origin: languages/c/tests/c/test_do_while.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n", "1\n"};
int __n = 2, __i = 0;
int i = 0;
do { if (i == 2) break; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } i++; } while (i < 5);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

