// vybe-test: c/pointers_arithmetic/pointer_decrement_from_one_past_end_reaches_last_element
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[3] = {4, 5, 6}; int *p = &arr[3];
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
p--;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

