// vybe-test: c/typedefs/typedef_alias_for_array_of_function_pointers_can_dispatch
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*Unary)(int); int add_one(int x) { return x + 1; } int double_it(int x) { return x * 2; }
int main() {
const char *__w[] = {"4 6\n"};
int __n = 1, __i = 0;
Unary ops[2] = {add_one, double_it}; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", ops[0](3), ops[1](3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

