// vybe-test: c/parameter_passing/pass_by_value_of_pointer_still_points_to_same_storage
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int read(int *p) { return *p; }
int main() {
const char *__w[] = {"13\n"};
int __n = 1, __i = 0;
int value = 13; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", read(&value));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

