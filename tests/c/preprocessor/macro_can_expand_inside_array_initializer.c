// vybe-test: c/preprocessor/macro_can_expand_inside_array_initializer
// origin: languages/c/tests/c/test_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define VALUE 7
int main() {
const char *__w[] = {"7 8\n"};
int __n = 1, __i = 0;
int arr[2] = {VALUE, VALUE + 1};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", arr[0], arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

