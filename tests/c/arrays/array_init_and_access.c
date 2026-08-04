// vybe-test: c/arrays/array_init_and_access
// origin: languages/c/tests/c/test_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"10\n", "50\n"};
int __n = 2, __i = 0;

    int arr[5] = {10, 20, 30, 40, 50};
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

