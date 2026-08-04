// vybe-test: c/restrict/restrict_pointer_parameter
// origin: languages/c/tests/c/test_restrict.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
void add_arrays(int n, int * restrict a, const int * restrict b) {
    for (int i = 0; i < n; i++) a[i] += b[i];
}
int main() {const char *__w[] = {"11 22 33\n"};
int __n = 1, __i = 0;

    int a[3] = {1, 2, 3};
    int b[3] = {10, 20, 30};
    add_arrays(3, a, b);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", a[0], a[1], a[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

