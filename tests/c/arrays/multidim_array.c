// vybe-test: c/arrays/multidim_array
// origin: languages/c/tests/c/test_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"2\n", "6\n"};
int __n = 2, __i = 0;

    int m[2][3] = {{1,2,3},{4,5,6}};
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", m[0][1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", m[1][2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

