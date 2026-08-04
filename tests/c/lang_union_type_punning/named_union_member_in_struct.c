// vybe-test: c/lang_union_type_punning/named_union_member_in_struct
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union Payload { int i; short s; }; struct Msg { union Payload p; };
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
struct Msg m; m.p.s = 7; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)m.p.s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

