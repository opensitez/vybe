// vybe-test: c/parameter_passing/array_parameter_pointer_arithmetic_reads_expected_slot
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int third(int *values) { return *(values + 2); }
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
int values[3] = {3, 6, 9}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", third(values));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

