// vybe-test: c/complex_initializers/union_init_first_member
// origin: languages/c/tests/c/test_complex_initializers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Val { int i; float f; };
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
union Val v = {42};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", v.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

