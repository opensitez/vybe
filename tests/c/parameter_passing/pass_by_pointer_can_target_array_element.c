// vybe-test: c/parameter_passing/pass_by_pointer_can_target_array_element
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void set_to_ten(int *x) { *x = 10; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int values[2] = {1, 2}; set_to_ten(&values[1]); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", values[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

