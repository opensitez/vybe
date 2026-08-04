// vybe-test: c/lang_storage_linkage/union_all_members_alias
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; float f; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
union U u; u.f=1.0f; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", u.i != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

