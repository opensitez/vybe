// vybe-test: c/c_restrict_pointers_overlap/restrict_pointer_overlap_fails
// origin: languages/c/tests/c/test_c_restrict_pointers_overlap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* void f(int *restrict p1, int *restrict p2) { *p1=1; *p2=2; } int main() { int x; f(&x, &x); return 0; } // UB to alias restrict pointers */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

