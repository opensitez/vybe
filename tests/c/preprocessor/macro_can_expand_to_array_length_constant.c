// vybe-test: c/preprocessor/macro_can_expand_to_array_length_constant
// origin: languages/c/tests/c/test_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define LEN 4
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int arr[LEN] = {1, 2, 3, 4};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[LEN - 1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

