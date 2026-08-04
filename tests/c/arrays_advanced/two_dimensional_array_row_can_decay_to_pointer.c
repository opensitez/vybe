// vybe-test: c/arrays_advanced/two_dimensional_array_row_can_decay_to_pointer
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int m[2][2] = {{2, 4}, {6, 8}};
int main() {
const char *__w[] = {"6 8\n"};
int __n = 1, __i = 0;
int *row = m[1];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", row[0], row[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

