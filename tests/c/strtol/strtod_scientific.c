// vybe-test: c/strtol/strtod_scientific
// origin: languages/c/tests/c/test_strtol.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"150.0\n"};
int __n = 1, __i = 0;
char *end;
double v = strtod("1.5e2", &end);
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

