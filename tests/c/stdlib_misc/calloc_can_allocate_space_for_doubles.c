// vybe-test: c/stdlib_misc/calloc_can_allocate_space_for_doubles
// origin: languages/c/tests/c/test_stdlib_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2.5\n"};
int __n = 1, __i = 0;
double *p = (double *)calloc(2, sizeof(double)); p[1] = 2.5; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", p[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

