// vybe-test: c/pointers_arithmetic/pointer_walk_back_and_forth_returns_to_same_value
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[4] = {10, 20, 30, 40}; int *p = &arr[1];
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;
p += 2;
p -= 2;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

