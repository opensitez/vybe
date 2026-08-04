// vybe-test: c/function_pointers/function_pointer_returning_pointer_can_be_called
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int *identity(int *p) { return p; }
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
int x = 7; int *(*fp)(int *) = identity;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *fp(&x));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

