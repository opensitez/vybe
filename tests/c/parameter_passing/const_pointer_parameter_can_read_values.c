// vybe-test: c/parameter_passing/const_pointer_parameter_can_read_values
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int first(const int *values) { return values[0]; }
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
int values[2] = {7, 8}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", first(values));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

