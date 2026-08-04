// vybe-test: c/stdlib_conversion_and_exit/strtol_endptr_at_suffix
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"e\n"};
int __n = 1, __i = 0;
char *e; strtol("42end", &e, 10); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", *e);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

