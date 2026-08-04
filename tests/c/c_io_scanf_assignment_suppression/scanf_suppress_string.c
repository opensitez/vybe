// vybe-test: c/c_io_scanf_assignment_suppression/scanf_suppress_string
// origin: languages/c/tests/c/test_c_io_scanf_assignment_suppression.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 world"};
int __n = 1, __i = 0;
 char buf[10]; int n = sscanf("hello world", "%*s %s", buf); { char __t[512]; snprintf(__t, sizeof(__t), "%d %s", n, buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

