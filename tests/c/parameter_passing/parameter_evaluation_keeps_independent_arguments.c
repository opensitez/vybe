// vybe-test: c/parameter_passing/parameter_evaluation_keeps_independent_arguments
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int combine(int a, int b) { return a * 10 + b; }
int main() {
const char *__w[] = {"12\n"};
int __n = 1, __i = 0;
int x = 1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", combine(x, x + 1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

