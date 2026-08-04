// vybe-test: c/parameter_passing/pointer_parameter_can_swap_two_values
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void swap(int *a, int *b) { int tmp = *a; *a = *b; *b = tmp; }
int main() {
const char *__w[] = {"2 1\n"};
int __n = 1, __i = 0;
int a = 1; int b = 2; swap(&a, &b); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

