// vybe-test: c/arrays_advanced/array_value_can_drive_for_loop_total
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[4] = {1, 3, 5, 7};
int main() {
const char *__w[] = {"16\n"};
int __n = 1, __i = 0;
int total = 0;
for (int i = 0; i < 4; i++) total += arr[i];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", total);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

