// vybe-test: c/string_io/snprintf_safe_truncate
// origin: languages/c/tests/c/test_string_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"hello w 11\n"};
int __n = 1, __i = 0;

char buf[8];
int n = snprintf(buf, sizeof(buf), "hello world");
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %d\n", buf, n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

