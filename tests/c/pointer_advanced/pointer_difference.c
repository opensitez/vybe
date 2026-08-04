// vybe-test: c/pointer_advanced/pointer_difference
// origin: languages/c/tests/c/test_pointer_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int arr[5];
int *a = &arr[1];
int *b = &arr[4];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)(b - a));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

