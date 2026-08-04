// vybe-test: c/lang_union_type_punning/union_nested_struct_member_write_read
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Pair { int a; int b; }; union U { struct Pair p; int i; };
int main() {
const char *__w[] = {"2 3\n"};
int __n = 1, __i = 0;
union U u; u.p.a = 2; u.p.b = 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", u.p.a, u.p.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

