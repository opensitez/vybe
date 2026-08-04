// vybe-test: c/loop_patterns/while_loop_digit_extraction
// origin: languages/c/tests/c/test_loop_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;

int n = 1234, count = 0;
while (n > 0) { count++; n /= 10; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

