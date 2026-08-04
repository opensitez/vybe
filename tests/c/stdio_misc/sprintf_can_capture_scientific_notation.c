// vybe-test: c/stdio_misc/sprintf_can_capture_scientific_notation
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char buf[64];
int main() {
const char *__w[] = {"1.2e+03\n"};
int __n = 1, __i = 0;
sprintf(buf, "%.1e", 1200.0);
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

