// vybe-test: c/lang_pointers_qualifiers/restrict_pointer_alias_hint
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int add(restrict int *a, restrict int *b) { return *a + *b; }
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int x=2,y=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", add(&x,&y));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

