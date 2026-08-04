// vybe-test: c/pointers_arithmetic/pointer_addition_inside_subscript_expression_reads_target
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[4] = {10, 20, 30, 40}; int *p = arr;
int main() {
const char *__w[] = {"30\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (p + 2)[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

