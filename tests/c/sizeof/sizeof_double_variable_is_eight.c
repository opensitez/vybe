// vybe-test: c/sizeof/sizeof_double_variable_is_eight
// origin: languages/c/tests/c/test_sizeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
double x = 7.0;
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)sizeof(x));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

