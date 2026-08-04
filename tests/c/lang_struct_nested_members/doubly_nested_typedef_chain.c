// vybe-test: c/lang_struct_nested_members/doubly_nested_typedef_chain
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
typedef struct { int v; } A; typedef struct { A a; } B; struct C { B b; };
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
struct C c = {{{7}}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", c.b.a.v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

