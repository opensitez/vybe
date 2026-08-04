// vybe-test: c/lang_struct_nested_members/struct_nested_in_union_outer_wrapper
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Inner { int n; }; union U { struct Inner s; int i; }; struct Wrap { union U u; };
int main() {
const char *__w[] = {"13\n"};
int __n = 1, __i = 0;
struct Wrap w; w.u.s.n = 13; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", w.u.s.n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

