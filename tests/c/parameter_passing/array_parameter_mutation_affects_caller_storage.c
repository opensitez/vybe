// vybe-test: c/parameter_passing/array_parameter_mutation_affects_caller_storage
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void set_second(int values[]) { values[1] = 9; }
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
int values[3] = {1, 2, 3}; set_second(values); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", values[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

