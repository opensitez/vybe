// vybe-test: c/functions_advanced/function_can_consume_pointer_to_pointer
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int read_twice(int **pp) { return **pp; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int x = 6; int *p = &x;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", read_twice(&p));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

