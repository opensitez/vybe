// vybe-test: c/c_stdio_formatted_io_edge_cases/scanf_eof_immediate
// origin: languages/c/tests/c/test_c_stdio_formatted_io_edge_cases.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"-1 9"};
int __n = 1, __i = 0;
 int x=9; int res = sscanf("", "%d", &x); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", res, x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

