// vybe-test: c/stdio_misc/printf_can_emit_multiple_lines_from_one_format
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"a\n", "b\n"};
int __n = 2, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "a\nb\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

