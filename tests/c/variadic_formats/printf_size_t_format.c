// vybe-test: c/variadic_formats/printf_size_t_format
// origin: languages/c/tests/c/test_variadic_formats.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
size_t s = sizeof(int);
{ char __t[512]; snprintf(__t, sizeof(__t), "%zu\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

