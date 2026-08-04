// vybe-test: c/lang_run_breadth3/nested_struct_copy
// origin: languages/c/tests/c/test_lang_run_breadth3.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct I{int v;}; struct O{struct I i;};
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct O a={{1}},b=a; b.i.v=2; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a.i.v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

