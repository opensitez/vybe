// vybe-test: c/multidim_arrays/two_dim_array_traversal
// origin: languages/c/tests/c/test_multidim_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n", "2\n", "3\n", "4\n"};
int __n = 4, __i = 0;

int m[2][2] = {{1,2},{3,4}};
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", m[i][j]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

