// vybe-test: c/function_pointers/typedef_like_function_pointer_usage_without_typedef_still_calls
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int square(int x) { return x * x; }
int main() {
const char *__w[] = {"25\n"};
int __n = 1, __i = 0;
int (*fp)(int) = square;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (*fp)(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

