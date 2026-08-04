// vybe-test: c/enum_operations/enum_in_array_bounds
// origin: languages/c/tests/c/test_enum_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum { SIZE = 5 };
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int arr[SIZE];
for (int i = 0; i < SIZE; i++) arr[i] = i;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[SIZE-1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

