// vybe-test: c/arrays_advanced/array_copy_via_loop_moves_all_values
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int src[3] = {3, 6, 9}; int dst[3] = {0, 0, 0};
int main() {
const char *__w[] = {"3 6 9\n"};
int __n = 1, __i = 0;
for (int i = 0; i < 3; i++) dst[i] = src[i];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", dst[0], dst[1], dst[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

