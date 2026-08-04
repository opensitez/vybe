// vybe-test: c/typedefs/typedef_alias_for_function_pointer_can_invoke_target
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*Unary)(int); int add_one(int x) { return x + 1; }
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
Unary fn = add_one; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fn(4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

