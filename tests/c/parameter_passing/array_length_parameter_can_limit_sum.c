// vybe-test: c/parameter_passing/array_length_parameter_can_limit_sum
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int sum(int values[], int len) { int total = 0; for (int i = 0; i < len; i++) total += values[i]; return total; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int values[4] = {1,2,3,4}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum(values, 3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

