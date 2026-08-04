// vybe-test: c/unions/union_pointer_can_be_reassigned_between_values
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; };
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
union Data first; union Data second; first.i = 1; second.i = 2; union Data *p = &first; p = &second; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

