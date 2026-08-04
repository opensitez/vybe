// vybe-test: c/lang_cast_expression_semantics/const_int_pointer_cast_mutation
// origin: languages/c/tests/c/test_lang_cast_expression_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"13\n"};
int __n = 1, __i = 0;
const int c = 12; int *p = (int *)&c; *p = 13; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

