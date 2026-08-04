// vybe-test: c/stdlib_misc/calloc_pointer_can_be_cast_to_int_pointer_and_summed
// origin: languages/c/tests/c/test_stdlib_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
int *p = (int *)calloc(3, sizeof(int)); p[1] = 4; p[2] = 5; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p[0] + p[1] + p[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

