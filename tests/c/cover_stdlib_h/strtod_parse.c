// vybe-test: c/cover_stdlib_h/strtod_parse
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"1.25\n"};
int __n = 1, __i = 0;
char *end; double v = strtod("1.25", &end); { char __t[512]; snprintf(__t, sizeof(__t), "%.2f\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

