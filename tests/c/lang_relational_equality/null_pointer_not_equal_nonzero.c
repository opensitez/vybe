// vybe-test: c/lang_relational_equality/null_pointer_not_equal_nonzero
// origin: languages/c/tests/c/test_lang_relational_equality.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
int x=0; int *p=&x; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p!=0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

