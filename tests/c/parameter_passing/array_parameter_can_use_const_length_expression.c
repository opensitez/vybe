// vybe-test: c/parameter_passing/array_parameter_can_use_const_length_expression
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int sum2(int values[], int len) { int total = 0; for (int i = 0; i < len; i++) total += values[i]; return total; }
int main() {
const char *__w[] = {"11\n"};
int __n = 1, __i = 0;
enum { LEN = 2 }; int values[LEN] = {5, 6}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum2(values, LEN));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

