// vybe-test: c/pointer_to_array/pointer_to_array_2d_row
// origin: languages/c/tests/c/test_pointer_to_array.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"5 8\n"};
int __n = 1, __i = 0;

int m[3][4] = {{1,2,3,4},{5,6,7,8},{9,10,11,12}};
int (*row)[4] = &m[1];
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", (*row)[0], (*row)[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

