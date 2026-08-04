// vybe-test: c/lang_pointers_qualifiers/pointer_to_const_value
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
const int x = 7; int *p = (int*)&x; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

