// vybe-test: c/parameter_passing/pointer_to_pointer_parameter_can_follow_two_levels
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int read2(int **pp) { return **pp; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int value = 6; int *p = &value; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", read2(&p));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

