// vybe-test: c/parameter_passing/double_pointer_parameter_can_retarget_pointer
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void point_to_second(int **pp, int *target) { *pp = target; }
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int a = 1; int b = 2; int *p = &a; point_to_second(&p, &b); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

