// vybe-test: c/function_pointers/function_pointer_and_function_symbol_can_share_same_behavior
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int triple(int x) { return x * 3; }
int main() {
const char *__w[] = {"9 9\n"};
int __n = 1, __i = 0;
int (*fp)(int) = triple;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", triple(3), fp(3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

