// vybe-test: c/c_for_loop_multiple_iterators/for_multiple_iter_pointers
// origin: languages/c/tests/c/test_c_for_loop_multiple_iterators.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"4"};
int __n = 1, __i = 0;
 int arr[] = {1,2,3}; int *p, *q; for(p=arr, q=arr+2; p<=q; p++, q--) *p += *q; { char __t[512]; snprintf(__t, sizeof(__t), "%d", arr[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

