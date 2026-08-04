// vybe-test: c/stdlib_misc/malloc_sizeof_array_pattern_can_store_multiple_ints
// origin: languages/c/tests/c/test_stdlib_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1 2 3\n"};
int __n = 1, __i = 0;
int *p = (int *)malloc(3 * sizeof(int)); p[0] = 1; p[1] = 2; p[2] = 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", p[0], p[1], p[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

