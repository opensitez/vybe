// vybe-test: c/function_pointers_advanced/fn_ptr_array_dispatch
// origin: languages/c/tests/c/test_function_pointers_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

int add(int a, int b) { return a + b; }
int sub(int a, int b) { return a - b; }
int mul(int a, int b) { return a * b; }
typedef int (*BinOp)(int, int);
int main() {
const char *__w[] = {"8 4 12\n"};
int __n = 1, __i = 0;

BinOp ops[3] = {add, sub, mul};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", ops[0](6,2), ops[1](6,2), ops[2](6,2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

