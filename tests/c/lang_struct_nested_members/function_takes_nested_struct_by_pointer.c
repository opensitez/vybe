// vybe-test: c/lang_struct_nested_members/function_takes_nested_struct_by_pointer
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Inner { int n; }; struct Outer { struct Inner in; }; void bump(struct Outer *o) { o->in.n++; }
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
struct Outer o = {{4}}; bump(&o); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", o.in.n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

