// vybe-test: c/lang_struct_nested_members/nested_struct_conditional_read
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Inner { int flag; int val; }; struct Outer { struct Inner in; };
int main() {
const char *__w[] = {"99\n"};
int __n = 1, __i = 0;
struct Outer o = {{1, 99}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", o.in.flag ? o.in.val : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

