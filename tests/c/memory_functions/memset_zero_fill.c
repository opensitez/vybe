// vybe-test: c/memory_functions/memset_zero_fill
// origin: languages/c/tests/c/test_memory_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0 0 0 0\n"};
int __n = 1, __i = 0;

int arr[4] = {1, 2, 3, 4};
memset(arr, 0, sizeof(arr));
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", arr[0], arr[1], arr[2], arr[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

