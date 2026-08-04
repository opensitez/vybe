// vybe-test: c/c_flexible_array_members_sizeof/fam_pass_by_value
// origin: languages/c/tests/c/test_c_flexible_array_members_sizeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"5"};
static int __n = 1, __i = 0;
struct S { int len; int data[]; }; void f(struct S s) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.len);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { struct S s = {5}; f(s); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

