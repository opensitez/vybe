// vybe-test: c/complex_typedefs/typedef_function_pointer
// origin: languages/c/tests/c/test_complex_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*BinaryOp)(int, int);
int add(int a, int b) { return a + b; }
int mul(int a, int b) { return a * b; }
int main() {
const char *__w[] = {"7 12\n"};
int __n = 1, __i = 0;
BinaryOp ops[2] = {add, mul};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", ops[0](3,4), ops[1](3,4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

